/**
 * GFM markdown → document block content. Parses the subset skill bodies and
 * AI drafts use (headings, paragraphs, bold/italic/strike, inline code, bullet
 * and ordered lists incl. nesting, pipe tables, blockquotes, links) into the
 * docx `Document` block model, so the markdown can be edited in a DOCX editor
 * and re-exported without drift.
 *
 * Round-trip notes:
 * - Lists are emitted as real list paragraphs (`listRendering` plus `numPr`),
 *   so an editor shows a marker and a markdown exporter re-derives `- ` /
 *   `1. ` rather than leaking a literal bullet glyph into the text.
 * - Inline code uses `Courier New`, a monospace family a markdown exporter can
 *   infer back to a backtick span.
 * - Every markdown list gets a matching `w:abstractNum`/`w:num` pair in the
 *   returned `numbering` (see {@link buildNumbering}), so the result is
 *   self-consistent and serialization never fails with a missing numbering
 *   definition. Merging this content onto a document that has its own
 *   numbering needs the two numbering namespaces renumbered apart.
 */
import type { Token, Tokens } from "marked";

import type {
  AbstractNumbering,
  BlockContent,
  ListLevel,
  ListRendering,
  NumberingDefinitions,
  NumberingInstance,
  Paragraph,
  ParagraphContent,
  Table,
  TableCell,
  TableRow,
} from "../model/document";
import { headingOutlineLevel } from "../model/outlineLevel";
import { paragraphNumberingReference } from "../model/paragraphNumbering";
import { sanitizeXmlCharacters } from "../serialize/xmlEscape";
import { inlineTokensToRuns, textRun } from "./inline";
import { isTokenType, lexMarkdown } from "./lexer";

/** Block content and the numbering definitions its list paragraphs reference. */
export type MarkdownContent = {
  content: BlockContent[];
  /** Present only when the markdown contained at least one list. */
  numbering?: NumberingDefinitions;
  /**
   * Non-fatal diagnostics from the compile, e.g. a table/code block/blockquote
   * inside a list item that the block model can't nest and had to keep as a
   * following sibling block instead. Present only when there is at least one.
   */
  warnings?: string[];
};

const para = (runs: ParagraphContent[], styleId?: string): Paragraph => ({
  type: "paragraph",
  formatting: styleId ? { styleId } : {},
  content: runs.length > 0 ? runs : [textRun("")],
});

const listPara = (runs: ParagraphContent[], rendering: ListRendering): Paragraph => ({
  type: "paragraph",
  // Real numbering properties, not just display metadata: an editor's list
  // commands (Enter continues the list, Tab indents, toggle) and the live
  // marker counters all key off `numPr`.
  formatting: {
    numPr: paragraphNumberingReference({ numId: rendering.numId, ilvl: rendering.level }),
  },
  listRendering: rendering,
  content: runs.length > 0 ? runs : [textRun("")],
});

// Header cells are not bolded: in GFM the header is positional (first row + the
// `---` separator), so bolding it would re-export as `**A**` and break the
// round-trip.
const cellOf = (cell: Tokens.TableCell): TableCell => ({
  type: "tableCell",
  content: [para(inlineTokensToRuns(cell.tokens, cell.text))],
});

const tableFromToken = (token: Tokens.Table): Table => ({
  type: "table",
  rows: [
    { type: "tableRow", cells: token.header.map((cell) => cellOf(cell)) },
    ...token.rows.map(
      (row): TableRow => ({
        type: "tableRow",
        cells: row.map((cell) => cellOf(cell)),
      }),
    ),
  ],
});

// Twips (720 = 0.5"). Each deeper level indents one more half-inch, matching
// the step the legal-source checklist profile uses for a single-column marker
// plus hanging indent.
const LIST_INDENT_STEP_TWIPS = 720;
const LIST_HANGING_INDENT_TWIPS = 360;

/**
 * One `w:abstractNum` level per (numId, ilvl) pair actually used by the
 * markdown, keyed by ilvl. Built alongside the blocks so the caller can
 * synthesize `numbering` afterwards: the DOCX serializer reads numbering
 * definitions only from there, never from the editor-only `listRendering`
 * hint.
 */
type NumIdLevels = Map<number, ListLevel>;

const buildListLevel = (ilvl: number, isBullet: boolean, start: number): ListLevel => ({
  ilvl,
  ...(!isBullet && { start }),
  numFmt: isBullet ? "bullet" : "decimal",
  lvlText: isBullet ? "•" : `%${ilvl + 1}.`,
  suffix: "tab",
  pPr: {
    indentLeft: LIST_INDENT_STEP_TWIPS * (ilvl + 1),
    indentFirstLine: -LIST_HANGING_INDENT_TWIPS,
    hangingIndent: true,
  },
});

// Real list paragraphs with Word-style template markers ("%1." resolves to the
// live counter at level 0), so inserted/split items renumber instead of
// repeating a baked-in number. Each top-level markdown list gets its own numId
// so separate lists restart at 1; nested lists share the parent's numId at a
// deeper ilvl. A markdown exporter resolves the templates back to concrete
// "N." markers and normalises bullets to "- ", so the markdown round-trips.
/**
 * The numId a list renders under. The first list to reach a (numId, ilvl)
 * pair defines that level; a later list at the same depth under the same
 * parent shares it when it is the same kind (so sibling nested bullets share
 * one counter), and gets a numId of its own when it is not (a nested ordered
 * list must not inherit a sibling's bullet definition).
 */
const resolveListNumId = (
  numIds: NumIdAllocator,
  parentNumId: number,
  level: ListLevel,
): number => {
  const levels = numIds.levels.get(parentNumId);
  const existing = levels?.get(level.ilvl);
  if (levels !== undefined && existing === undefined) {
    levels.set(level.ilvl, level);
    return parentNumId;
  }
  if (
    levels !== undefined &&
    existing !== undefined &&
    existing.numFmt === level.numFmt &&
    existing.start === level.start
  ) {
    return parentNumId;
  }
  const numId = numIds.next++;
  numIds.levels.set(numId, new Map([[level.ilvl, level]]));
  return numId;
};

/**
 * Whether an item-content token is one the block model cannot nest inside a
 * list item's own paragraph: a nested list (handled separately, at a deeper
 * `ilvl`) or a block the DOCX list model has no notion of nesting at all
 * (OOXML lists are just numbered paragraphs, not containers — a table,
 * code block, or blockquote next to one is always a sibling, never a child).
 */
const isNestedBlockToken = (token: Token | undefined): boolean =>
  token !== undefined &&
  (isTokenType(token, "list") ||
    isTokenType(token, "table") ||
    isTokenType(token, "code") ||
    isTokenType(token, "blockquote"));

const listBlocks = (
  list: Tokens.List,
  level: number,
  parentNumId: number,
  state: CompileState,
): BlockContent[] => {
  const out: BlockContent[] = [];
  const start = Number(list.start) || 1;
  const decimalLevels = Array.from({ length: level + 1 }, () => "decimal" as const);
  const numId = resolveListNumId(state, parentNumId, buildListLevel(level, !list.ordered, start));
  for (const item of list.items) {
    const rendering: ListRendering = list.ordered
      ? {
          marker: `%${level + 1}.`,
          level,
          numId,
          isBullet: false,
          numFmt: "decimal",
          levelNumFmts: decimalLevels,
          ...(start !== 1 && { startOverride: start }),
        }
      : { marker: "•", level, numId, isBullet: true };
    // The item's own paragraph is every leading token up to the first block
    // the model can't fold into it; everything from there on (that block, and
    // anything after it) is handled separately below, in source order. The
    // blank-line `space` token right before that first block is a pure
    // separator, not a continuation break inside the item's own paragraph
    // (unlike a `space` between two inline paragraphs of a loose item, which
    // does belong there — see `inline.ts`'s handling of it), so it is dropped
    // rather than folded into either side.
    const inlineTokens: Token[] = [];
    const trailing: Token[] = [];
    let inTrailing = false;
    for (const [index, child] of item.tokens.entries()) {
      if (!inTrailing && isNestedBlockToken(child)) {
        inTrailing = true;
      }
      if (!inTrailing && child.type === "space" && isNestedBlockToken(item.tokens[index + 1])) {
        continue;
      }
      if (inTrailing) {
        trailing.push(child);
      } else {
        inlineTokens.push(child);
      }
    }
    out.push(listPara(inlineTokensToRuns(inlineTokens, item.text), rendering));
    for (const block of trailing) {
      if (isTokenType(block, "list")) {
        // A true nested list: a deeper `ilvl` under the same item, the one
        // case the model can actually express as nesting.
        out.push(...listBlocks(block, level + 1, numId, state));
        continue;
      }
      if (block.type === "space") {
        // Pure whitespace between the item's text and its trailing block(s);
        // nothing to preserve.
        continue;
      }
      // A table, code block, blockquote, or (rarer) further prose that
      // followed one of those inside the item: keep it as a block right
      // after the item instead of silently dropping it, since the model has
      // no way to nest it inside the item's own paragraph.
      const rendered = blocksFromTokens([block], state);
      if (rendered.length > 0) {
        out.push(...rendered);
        state.warnings.push(
          `A list item's ${block.type} content could not stay nested inside the item ` +
            "(the DOCX list model has no way to nest it); it was kept as a block right " +
            "after the item instead.",
        );
      }
    }
  }
  return out;
};

/**
 * Allocates one numId per markdown list so each list counts independently,
 * and collects the level definitions needed to synthesize `numbering` for
 * every list it mints. Also carries diagnostics collected while walking the
 * tokens (list item content the block model can't nest — see `listBlocks`).
 */
type NumIdAllocator = { next: number; levels: Map<number, NumIdLevels> };
type CompileState = NumIdAllocator & { warnings: string[] };

/**
 * The deepest heading the built-in style sets define, so a `#####` compiles to
 * a style the document actually has.
 */
const MAX_HEADING_LEVEL = 4;

/**
 * The style id for a markdown heading, plus the outline level it stands for.
 *
 * The id is English because the document this content lands in is one folio
 * created, and folio's style sets define `Heading1`…`HeadingN` under the
 * built-in name. The outline level is what makes the classification survive
 * anyway: a consumer reads `w:outlineLvl` first (see
 * `@stll/folio-core/docx/builtInStyles`), so the paragraph stays a heading
 * even when merged into a document whose own heading styles are localized and
 * this id resolves to nothing.
 */
const headingParagraph = (runs: ParagraphContent[], depth: number): Paragraph => {
  const level = Math.min(Math.max(depth, 1), MAX_HEADING_LEVEL);
  const paragraph = para(runs, `Heading${level}`);
  const outlineLevel = headingOutlineLevel(level - 1);
  if (outlineLevel !== undefined) {
    paragraph.formatting = { ...paragraph.formatting, outlineLevel };
  }
  return paragraph;
};

/** The built-in a markdown blockquote compiles to. */
const QUOTE_STYLE_ID = "Quote";

const blocksFromTokens = (tokens: Token[] | undefined, state: CompileState): BlockContent[] => {
  const blocks: BlockContent[] = [];
  for (const token of tokens ?? []) {
    if (isTokenType(token, "heading")) {
      blocks.push(headingParagraph(inlineTokensToRuns(token.tokens, token.text), token.depth));
    } else if (isTokenType(token, "paragraph")) {
      blocks.push(para(inlineTokensToRuns(token.tokens, token.text)));
    } else if (isTokenType(token, "list")) {
      const numId = state.next++;
      state.levels.set(numId, new Map());
      blocks.push(...listBlocks(token, 0, numId, state));
    } else if (isTokenType(token, "table")) {
      blocks.push(tableFromToken(token));
    } else if (isTokenType(token, "code")) {
      for (const line of token.text.split("\n")) {
        blocks.push(para([textRun(line.length > 0 ? line : " ", { mono: true })]));
      }
    } else if (isTokenType(token, "blockquote")) {
      for (const inner of blocksFromTokens(token.tokens, state)) {
        const styled: BlockContent =
          inner.type === "paragraph"
            ? {
                ...inner,
                formatting: { ...inner.formatting, styleId: QUOTE_STYLE_ID },
              }
            : inner;
        blocks.push(styled);
      }
    } else if (token.type === "hr") {
      blocks.push(para([textRun("———")]));
    } else if (
      token.type !== "space" &&
      "text" in token &&
      typeof token.text === "string" &&
      token.text.trim().length > 0
    ) {
      blocks.push(para([textRun(token.text)]));
    }
  }
  return blocks;
};

// One `w:abstractNum` per numId (a 1:1 mapping, so `abstractNumId === numId`
// keeps the synthesis trivial to reason about; callers merging this into a
// document with its own numbering should not assume the mapping stays 1:1
// after remapping). Every level actually visited by that markdown list
// becomes one `w:lvl`, so serializing the content never references a numId
// with no definition.
const buildNumbering = (numIdLevels: Map<number, NumIdLevels>): NumberingDefinitions => {
  const abstractNums: AbstractNumbering[] = [];
  const nums: NumberingInstance[] = [];
  for (const [numId, levels] of numIdLevels) {
    const sortedLevels = [...levels.entries()].sort(([a], [b]) => a - b).map(([, lvl]) => lvl);
    abstractNums.push({
      abstractNumId: numId,
      multiLevelType: sortedLevels.length > 1 ? "multilevel" : "singleLevel",
      levels: sortedLevels,
    });
    nums.push({ numId, abstractNumId: numId });
  }
  return { abstractNums, nums };
};

/**
 * Parse GFM markdown into document blocks plus the numbering its lists need.
 * Synchronous. The caller places the blocks into a `Document` of its own
 * (page geometry, styles, and presets are the host's decision).
 */
export const compileMarkdownToContent = (markdown: string): MarkdownContent => {
  const state: CompileState = { next: 1, levels: new Map(), warnings: [] };
  // Markdown arrives as text from outside folio, so it can carry characters no
  // XML document may contain. They are dropped here, while the input is still
  // one string, rather than at the serializer, where nothing could say which
  // input lost them. This particular drop has no channel to report it on.
  const content = blocksFromTokens(lexMarkdown(sanitizeXmlCharacters(markdown)), state);
  return {
    content,
    ...(state.levels.size > 0 && { numbering: buildNumbering(state.levels) }),
    ...(state.warnings.length > 0 && { warnings: state.warnings }),
  };
};
