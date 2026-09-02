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
import { inlineTokensToRuns, textRun } from "./inline";
import { isTokenType, lexMarkdown } from "./lexer";

/** Block content and the numbering definitions its list paragraphs reference. */
export type MarkdownContent = {
  content: BlockContent[];
  /** Present only when the markdown contained at least one list. */
  numbering?: NumberingDefinitions;
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
  formatting: { numPr: { numId: rendering.numId, ilvl: rendering.level } },
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

const listBlocks = (
  list: Tokens.List,
  level: number,
  parentNumId: number,
  numIds: NumIdAllocator,
): BlockContent[] => {
  const out: BlockContent[] = [];
  const start = Number(list.start) || 1;
  const decimalLevels = Array.from({ length: level + 1 }, () => "decimal" as const);
  const numId = resolveListNumId(numIds, parentNumId, buildListLevel(level, !list.ordered, start));
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
    const inlineTokens: Token[] = [];
    const nestedLists: Tokens.List[] = [];
    for (const child of item.tokens) {
      if (isTokenType(child, "list")) {
        nestedLists.push(child);
      } else {
        inlineTokens.push(child);
      }
    }
    out.push(listPara(inlineTokensToRuns(inlineTokens, item.text), rendering));
    for (const nested of nestedLists) {
      out.push(...listBlocks(nested, level + 1, numId, numIds));
    }
  }
  return out;
};

/**
 * Allocates one numId per markdown list so each list counts independently,
 * and collects the level definitions needed to synthesize `numbering` for
 * every list it mints.
 */
type NumIdAllocator = { next: number; levels: Map<number, NumIdLevels> };

const MAX_HEADING_LEVEL = 4;

const blocksFromTokens = (tokens: Token[] | undefined, numIds: NumIdAllocator): BlockContent[] => {
  const blocks: BlockContent[] = [];
  for (const token of tokens ?? []) {
    if (isTokenType(token, "heading")) {
      const level = Math.min(Math.max(token.depth, 1), MAX_HEADING_LEVEL);
      blocks.push(para(inlineTokensToRuns(token.tokens, token.text), `Heading${level}`));
    } else if (isTokenType(token, "paragraph")) {
      blocks.push(para(inlineTokensToRuns(token.tokens, token.text)));
    } else if (isTokenType(token, "list")) {
      const numId = numIds.next++;
      numIds.levels.set(numId, new Map());
      blocks.push(...listBlocks(token, 0, numId, numIds));
    } else if (isTokenType(token, "table")) {
      blocks.push(tableFromToken(token));
    } else if (isTokenType(token, "code")) {
      for (const line of token.text.split("\n")) {
        blocks.push(para([textRun(line.length > 0 ? line : " ", { mono: true })]));
      }
    } else if (isTokenType(token, "blockquote")) {
      for (const inner of blocksFromTokens(token.tokens, numIds)) {
        const styled: BlockContent =
          inner.type === "paragraph"
            ? {
                ...inner,
                formatting: { ...inner.formatting, styleId: "Quote" },
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
  const numIds: NumIdAllocator = { next: 1, levels: new Map() };
  const content = blocksFromTokens(lexMarkdown(markdown), numIds);
  return {
    content,
    ...(numIds.levels.size > 0 && { numbering: buildNumbering(numIds.levels) }),
  };
};
