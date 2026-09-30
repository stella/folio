/**
 * Render inline content (runs, hyperlinks, comment-range markers, tracked
 * changes) to markdown. Operates on the `ParagraphContent[]` of a paragraph, or
 * the equivalent inline lists inside hyperlinks and tracked-change wrappers.
 *
 * Inline marks are rendered as character-precise wrappers around the affected
 * runs. Comments and tracked changes become configurable annotation tags via
 * `./annotations`. Ported from eigenpal/docx-editor PR #595 (folio adds the
 * `footnotes: "strip"` gate).
 */

import type {
  Comment,
  CommentRangeEnd,
  CommentRangeStart,
  Deletion,
  DocxPackage,
  Hyperlink,
  Insertion,
  MoveFrom,
  MoveTo,
  ParagraphContent,
  Run,
  RunContent,
  TrackedRunContent,
} from "../types/document";
import { compileMarkdownToContent } from "@stll/docx-core";

import { decodeOoxmlSymbolCharacter } from "../utils/ooxmlSymbol";
import { wrapComment, wrapDeletion, wrapInsertion, wrapMoveFrom, wrapMoveTo } from "./annotations";
import { escapeAltText, escapeInline, escapeLinkUrl } from "./escape";
import { registerImage } from "./images";
import { getHyperlinkRuns } from "../docx/hyperlinkParser";
import { RELATIONSHIP_TYPES, resolveRelationshipIdOfType } from "../docx/relsParser";
import { numberNoteReference, pushWarning } from "./internals";
import type { RenderContext } from "./types";

/**
 * Inline marks we recognize. Order matters: we open from outermost to innermost
 * so the output reads cleanly (`***bold italic***`), not the reverse.
 */
type MarkKey = "bold" | "italic" | "code" | "strike";

const MARK_DELIMS: Record<MarkKey, string> = {
  bold: "**",
  italic: "*",
  code: "`",
  strike: "~~",
};

// Word has no `code` run property. We infer it from a small whitelist of
// monospace font families so prose set in fonts like `Monotype Corsiva` is not
// wrapped in backticks.
const MONOSPACE_FONTS = new Set([
  "consolas",
  "courier",
  "courier new",
  "menlo",
  "monaco",
  "sf mono",
  "jetbrains mono",
  "fira code",
  "fira mono",
  "source code pro",
  "roboto mono",
  "inconsolata",
  "lucida console",
  "monospace",
]);

function marksFor(run: Run): MarkKey[] {
  const f = run.formatting;
  if (!f) {
    return [];
  }
  const out: MarkKey[] = [];
  // Innermost first. Strikethrough goes inside emphasis: an emphasis
  // delimiter beside a tilde opens and closes wherever it stands, where the
  // other way round (`~~*`) needs a blank or punctuation outside.
  if (f.strike) {
    out.push("strike");
  }
  if (f.bold) {
    out.push("bold");
  }
  if (f.italic) {
    out.push("italic");
  }
  const ascii = f.fontFamily?.ascii?.toLowerCase();
  if (ascii && MONOSPACE_FONTS.has(ascii)) {
    out.push("code");
  }
  return out;
}

function applyMarks(text: string, marks: MarkKey[]): string {
  if (!text) {
    return text;
  }
  // Code overrides other marks: a code span is literal, so undo the inline
  // markdown escaping `escapeInline` applied to the run text. Per CommonMark the
  // fence must be longer than the longest backtick run in the content, and a
  // space is padded inside when the content begins or ends with a backtick.
  if (marks.includes("code")) {
    const literal = text.replace(/\\(?<char>[\\`*[\]_<~])/gu, "$<char>");
    let longestRun = 0;
    for (const m of literal.matchAll(/`+/gu)) {
      longestRun = Math.max(longestRun, m[0].length);
    }
    const fence = "`".repeat(longestRun + 1);
    const pad = literal.startsWith("`") || literal.endsWith("`") ? " " : "";
    return `${fence}${pad}${literal}${pad}${fence}`;
  }
  let out = text;
  for (const m of marks) {
    const d = MARK_DELIMS[m];
    out = `${d}${out}${d}`;
  }
  return out;
}

/**
 * The character beside a run's delimiters: `""` at the paragraph's edge,
 * `undefined` when it cannot be told (read as a letter, the strictest case).
 */
type Neighbor = string | undefined;

type RunNeighbors = { before: Neighbor; after: Neighbor };

const UNKNOWN_NEIGHBORS: RunNeighbors = { before: undefined, after: undefined };

const isBlank = (ch: Neighbor): boolean => ch === "" || (ch !== undefined && /\s/u.test(ch));

type EmphasisKey = Exclude<MarkKey, "code">;

/**
 * Which mark wraps a stretch of runs first, outermost first. Strikethrough
 * goes inside emphasis: an emphasis delimiter beside a tilde opens and closes
 * wherever it stands, where the other way round (`~~*`) needs a blank or
 * punctuation outside.
 */
const GROUPING_ORDER: readonly EmphasisKey[] = ["italic", "bold", "strike"];

/** The stackings tried, in turn, where the reader misreads one. */
const GROUPING_ORDERS: readonly (readonly EmphasisKey[])[] = [
  GROUPING_ORDER,
  ["bold", "italic", "strike"],
];

/**
 * CommonMark punctuation (Unicode P and S categories) as a delimiter of
 * `mark` meets it. Beside an emphasis delimiter, an asterisk is another
 * delimiter that runs together with it, and a tilde does not let it close
 * (the GFM reader's rule): both count as letters. (Words carry them escaped,
 * as a two-character unit.)
 */
const punctuationFor =
  (mark: EmphasisKey) =>
  (ch: Neighbor): boolean => {
    if (ch === undefined || ch === "" || ch === "~") return false;
    if (mark !== "strike" && ch === "*") return false;
    return /[\p{P}\p{S}]/u.test(ch);
  };

/** One escape pair (`\*`) or one code point of escaped inline text. */
const INLINE_UNIT = /\\[\s\S]|[\s\S]/gu;

/**
 * Wrap `inner` in `mark` so a CommonMark reader reads it back. A delimiter
 * opens only before a character that is not blank, and before punctuation
 * only where blank or punctuation precedes it; closing mirrors that. Where
 * the words at an edge would stop a delimiter from opening or closing (a
 * space at the edge, emphasis ending in a full stop right before a letter),
 * those characters are left outside the emphasis rather than written as
 * literal asterisks.
 */
function wrapFlanking(inner: string, mark: EmphasisKey, { before, after }: RunNeighbors): string {
  const isPunctuation = punctuationFor(mark);
  const opensAfter = (ch: Neighbor) => isBlank(ch) || isPunctuation(ch);
  const units = Array.from(inner.matchAll(INLINE_UNIT), ([unit]) => unit);
  const isPunctuationUnit = (unit: string) =>
    (unit.length === 2 && unit.startsWith("\\")) || isPunctuation(unit);
  let start = 0;
  let end = units.length;
  let outside = before;
  for (;;) {
    const unit = units[start];
    if (unit === undefined || start >= end) break;
    const opens = !isBlank(unit) && (!isPunctuationUnit(unit) || opensAfter(outside));
    if (opens) break;
    outside = unit.at(-1);
    start++;
  }
  outside = after;
  for (;;) {
    const unit = units[end - 1];
    if (unit === undefined || end <= start) break;
    const closes = !isBlank(unit) && (!isPunctuationUnit(unit) || opensAfter(outside));
    if (closes) break;
    outside = unit.at(0);
    end--;
  }
  return (
    units.slice(0, start).join("") +
    applyMarks(units.slice(start, end).join(""), start < end ? [mark] : []) +
    units.slice(end).join("")
  );
}

/** Paragraph content that writes nothing to the line. */
const SILENT_MARKERS = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

/** Comment range markers, which write nothing when comments are stripped. */
const COMMENT_MARKERS = new Set(["commentRangeStart", "commentRangeEnd"]);

/** A run as rendered text, with the emphasis still to wrap around it. */
type Piece = { text: string; marks: ReadonlySet<EmphasisKey> };

/** The first character `piece` puts on the line, inside the marks `outer` already opened. */
const pieceStart = (
  piece: Piece,
  outer: ReadonlySet<EmphasisKey>,
  order: readonly EmphasisKey[],
): Neighbor => {
  const first = piece.text.match(INLINE_UNIT)?.[0];
  if (first === undefined || isBlank(first)) return first;
  const opens = order.find((mark) => piece.marks.has(mark) && !outer.has(mark));
  if (opens === undefined) return first;
  return opens === "strike" ? "~" : "*";
};

/**
 * Wrap consecutive pieces in their emphasis, one mark at a time: the pieces
 * sharing the outermost mark are wrapped once, around the rest of their
 * emphasis. Written run by run, the closing delimiter of one would meet the
 * opening one of the next (`*a**b*`), which no reader takes for two
 * emphases.
 */
function renderPieces(
  pieces: readonly Piece[],
  outer: ReadonlySet<EmphasisKey>,
  edges: RunNeighbors,
  order: readonly EmphasisKey[] = GROUPING_ORDER,
): string {
  let out = "";
  let index = 0;
  while (index < pieces.length) {
    const piece = pieces[index];
    if (!piece) break;
    const mark = order.find((key) => piece.marks.has(key) && !outer.has(key));
    if (mark === undefined) {
      out += piece.text;
      index++;
      continue;
    }
    let last = index;
    while (pieces[last + 1]?.marks.has(mark)) last++;
    const next = pieces[last + 1];
    // A delimiter opened right inside this one runs together with it into
    // one delimiter run, which opens or closes by what stands outside both.
    const around: RunNeighbors = {
      before: out.length > 0 ? out.at(-1) : edges.before,
      after: next === undefined ? edges.after : pieceStart(next, outer, order),
    };
    const inner = renderPieces(
      pieces.slice(index, last + 1),
      new Set([...outer, mark]),
      around,
      order,
    );
    out += wrapFlanking(inner, mark, around);
    index = last + 1;
  }
  return out;
}

/** The first character an item that is not a run puts on the line, where it is fixed. */
function renderedStart(item: { type: string } | undefined): Neighbor {
  if (item?.type !== "hyperlink") return undefined;
  const link = item as unknown as Hyperlink;
  return link.href || link.anchor ? "[" : undefined;
}

/** Render a single run's RunContent array into the inline text fragment. */
function renderRunContent(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  content: RunContent[],
  paraId: string | undefined,
): string {
  let out = "";
  for (const item of content) {
    switch (item.type) {
      case "text":
        out += escapeInline(item.text);
        break;
      case "tab":
        out += "    ";
        break;
      case "break":
        if (item.breakType === "page") {
          // Page break inside text; in unpaged output we emit a paragraph break.
          out += "\n\n";
        } else {
          // Soft break: markdown's two-space hard wrap.
          out += "  \n";
        }
        break;
      case "symbol":
        out += escapeInline(decodeOoxmlSymbolCharacter(item.char) ?? item.char);
        break;
      case "softHyphen":
        // U+00AD soft hyphen. Word displays it only when needed for line
        // breaks; drop it from the markdown output.
        break;
      case "noBreakHyphen":
        out += "‑";
        break;
      // Opaque markup contributes only what it puts on the line: a `w:ruby`
      // base is a word in the sentence, the rest shows nothing.
      case "preservedXml":
        out += escapeInline(item.text);
        break;
      case "footnoteRef":
      case "endnoteRef": {
        if (ctx.opts.footnotes === "strip") {
          break;
        }
        out += numberNoteReference(
          ctx,
          item.type === "endnoteRef" ? "endnote" : "footnote",
          item.id,
        ).marker;
        break;
      }
      case "drawing": {
        // Preferred path: resolve via the package's rels → media chain. That
        // returns raw bytes, so we register a stable virtual path and expose
        // the image in `result.images`.
        const ref = resolveRelationshipIdOfType(
          pkg?.relationships,
          item.image.rId,
          RELATIONSHIP_TYPES.image,
        );
        const media =
          ref.status === "resolved" ? pkg?.media?.get(ref.relationship.target) : undefined;
        if (media) {
          const reg = registerImage(ctx, media, item.image, paraId);
          const alt = reg.alt ? escapeAltText(reg.alt) : "";
          out += `![${alt}](${reg.virtualPath})`;
          break;
        }
        // Fallback: header/footer images use a separate rels file that does not
        // live in `pkg.relationships`. The parser inlines the bytes into
        // `image.src` (typically a data URL) — emit that directly.
        if (item.image.src) {
          const alt = item.image.alt ?? item.image.title ?? item.image.filename ?? "";
          out += `![${escapeAltText(alt)}](${item.image.src})`;
          break;
        }
        pushWarning(ctx, `image rId=${item.image.rId ?? "(absent)"} not resolvable`);
        break;
      }
      case "shape":
        pushWarning(ctx, "shape not representable in markdown");
        break;
      case "fieldChar":
      case "instrText":
        // Field chrome. Skip: the field result text lives in surrounding runs.
        break;
      default:
        break;
    }
  }
  return out;
}

/** A run's rendered text and the emphasis it still takes; `null` when it writes nothing. */
function pieceOf(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  run: Run,
  paraId: string | undefined,
): Piece | null {
  // Hidden text (`w:vanish`) is suppressed in Word's normal view; drop it from
  // the markdown so hidden clauses / drafting notes don't leak into the export.
  if (run.formatting?.hidden) {
    return null;
  }
  const inner = renderRunContent(ctx, pkg, run.content, paraId);
  if (!inner) {
    return null;
  }
  const marks = marksFor(run);
  if (!marks.includes("code")) {
    return { text: inner, marks: new Set(marks as EmphasisKey[]) };
  }
  // A code span cannot hold whitespace at its ends, so it is split out.
  // Done with trim-length math rather than a regex to avoid backtracking on
  // long runs.
  const leadLen = inner.length - inner.trimStart().length;
  const trailLen = inner.length - inner.trimEnd().length;
  const core = inner.slice(leadLen, inner.length - trailLen);
  const text = core
    ? `${inner.slice(0, leadLen)}${applyMarks(core, marks)}${inner.slice(inner.length - trailLen)}`
    : inner;
  return { text, marks: new Set() };
}

type ReadCharacter = { char: string; marks: ReadonlySet<EmphasisKey> };

/** What a markdown reader reads from one line: each character with its emphasis. */
const readLine = (line: string): ReadCharacter[] =>
  compileMarkdownToContent(line)
    .content.flatMap((block) => (block.type === "paragraph" ? block.content : []))
    .flatMap((item) => (item.type === "hyperlink" ? item.children : [item]))
    .flatMap((item) => {
      if (item.type !== "run") return [{ char: "\u0000", marks: new Set<EmphasisKey>() }];
      const marks = new Set<EmphasisKey>(
        GROUPING_ORDER.filter((key) => item.formatting?.[key] === true),
      );
      return item.content.flatMap((content) =>
        content.type === "text"
          ? Array.from(content.text, (char) => ({ char, marks }))
          : [{ char: "\u0000", marks }],
      );
    });

const sameMarks = (left: ReadonlySet<EmphasisKey>, right: ReadonlySet<EmphasisKey>) =>
  left.size === right.size && [...left].every((key) => right.has(key));

/**
 * How `rendered`, between `edges`, reads back: `"exact"` with the pieces'
 * text and each letter's emphasis, `"text"` with their text alone (emphasis
 * lost somewhere, none of it read as literal asterisks), or `null`.
 */
const readBack = (
  rendered: string,
  pieces: readonly Piece[],
  edges: RunNeighbors,
): "exact" | "text" | null => {
  // A paragraph's edge reads like a blank; an unknown neighbour like a letter.
  const side = (ch: Neighbor): string => {
    if (ch === undefined) return "x";
    return ch === "" ? " " : ch;
  };
  const line = (inner: string) => `x ${side(edges.before)}${inner}${side(edges.after)} x`;
  const read = readLine(line(rendered));
  const plain = readLine(line(pieces.map(({ text }) => text).join("")));
  if (read.map(({ char }) => char).join("") !== plain.map(({ char }) => char).join("")) {
    return null;
  }
  // What the reader reads from each piece on its own (a hard break or an
  // image reads as one item), with the emphasis the piece should carry.
  const expected = pieces.flatMap(({ text, marks }) =>
    readLine(`x${text}x`)
      .slice(1, -1)
      .map(({ char }) => ({ char, marks })),
  );
  const start = Array.from(`x ${side(edges.before)}`).length;
  const end = Array.from(`${side(edges.after)} x`).length;
  if (read.length !== start + expected.length + end) {
    return read.every(({ marks }) => marks.size === 0) ? "text" : null;
  }
  let exact = true;
  for (const [index, { char, marks }] of expected.entries()) {
    const at = read[start + index];
    if (at?.char !== char) return null;
    // Emphasis the words did not have is never written.
    if ([...at.marks].some((mark) => !marks.has(mark))) return null;
    if (/[\p{L}\p{N}]/u.test(char) && !sameMarks(at.marks, marks)) exact = false;
  }
  return exact ? "exact" : "text";
};

/**
 * Render consecutive runs with their formatting applied. The delimiters are
 * placed by the flanking rules the reader applies, and the result is checked
 * against the reader: a stacking it misreads is tried the other way round,
 * or with more punctuation left outside, and failing all of those, the runs
 * go out without emphasis rather than with stray asterisks.
 */
function renderRuns(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  runs: readonly Run[],
  paraId: string | undefined,
  edges: RunNeighbors = UNKNOWN_NEIGHBORS,
): string {
  const pieces = runs.flatMap((run) => pieceOf(ctx, pkg, run, paraId) ?? []);
  const plain = pieces.map(({ text }) => text).join("");
  if (!pieces.some(({ marks }) => marks.size > 0)) {
    return plain;
  }
  let fallback: string | undefined;
  // Emphasis spanning a space can nest past what the reader matches; the
  // last tries close it at each space instead, leaving the spaces plain.
  for (const split of [pieces, splitAtBlanks(pieces)]) {
    for (const order of GROUPING_ORDERS) {
      for (const neighbors of [edges, UNKNOWN_NEIGHBORS]) {
        const rendered = renderPieces(split, new Set(), neighbors, order);
        const read = readBack(rendered, pieces, edges);
        if (read === "exact") return rendered;
        if (read === "text") fallback ??= rendered;
      }
    }
  }
  return fallback ?? plain;
}

/** The pieces with every blank a piece of its own, without emphasis. */
const splitAtBlanks = (pieces: readonly Piece[]): Piece[] =>
  pieces.flatMap((piece) => {
    if (piece.marks.size === 0) return [piece];
    const parts: Piece[] = [];
    for (const [unit] of piece.text.matchAll(INLINE_UNIT)) {
      const marks = isBlank(unit) ? new Set<EmphasisKey>() : piece.marks;
      const last = parts.at(-1);
      if (last && last.marks === marks) last.text += unit;
      else parts.push({ text: unit, marks });
    }
    return parts;
  });

function renderHyperlink(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  link: Hyperlink,
  paraId: string | undefined,
): string {
  // Through a transparent wrapper: markdown carries neither a layout control
  // nor a tag name, so the linked text inside one is still the link's text.
  const bracketed = Boolean(link.href ?? link.anchor);
  const inner = renderRuns(
    ctx,
    pkg,
    getHyperlinkRuns(link),
    paraId,
    bracketed ? { before: "[", after: "]" } : UNKNOWN_NEIGHBORS,
  );
  if (!inner) {
    return "";
  }
  const href = link.href ?? (link.anchor ? `#${link.anchor}` : "");
  if (!href) {
    pushWarning(ctx, "hyperlink missing href and anchor; rendered as plain text");
    return inner;
  }
  if (ctx.opts.hyperlinks === "reference") {
    const refNumber = ctx.hyperlinkRefs.length + 1;
    ctx.hyperlinkRefs.push({ href, refNumber });
    return `[${inner}][${refNumber}]`;
  }
  return `[${inner}](${escapeLinkUrl(href)})`;
}

function renderTrackedWrapper(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  wrapper: Insertion | Deletion | MoveFrom | MoveTo,
  paraId: string | undefined,
): string {
  const renderChild = (child: TrackedRunContent): string => {
    switch (child.type) {
      case "run":
        return renderRuns(ctx, pkg, [child], paraId);
      case "hyperlink":
        return renderHyperlink(ctx, pkg, child, paraId);
      case "mathEquation":
        return child.plainText ? escapeInline(child.plainText) : "";
      // A transparent wrapper carries the revision's text; reading through it
      // is the only way that text reaches the output.
      case "inlineWrapper":
      case "inlineSdt":
        return renderParagraphInline(ctx, pkg, child.content, paraId, UNKNOWN_NEIGHBORS);
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        return renderTrackedWrapper(ctx, pkg, child, paraId);
      case "simpleField":
      case "complexField":
        return renderParagraphInline(ctx, pkg, [child], paraId, UNKNOWN_NEIGHBORS);
      // A range boundary carries no text, and neither does markup folio kept
      // as bytes: a capture is not read, so it has no words to render.
      case "bookmarkStart":
      case "bookmarkEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "preservedInline":
        return "";
      default: {
        const unrendered: never = child;
        return unrendered;
      }
    }
  };
  // Consecutive runs are written together, so their emphasis joins up.
  const renderChildren = (): string => {
    let rendered = "";
    let runs: Run[] = [];
    const flush = (after: Neighbor) => {
      if (runs.length === 0) return;
      rendered += renderRuns(ctx, pkg, runs, paraId, {
        before: rendered.length > 0 ? rendered.at(-1) : undefined,
        after,
      });
      runs = [];
    };
    for (const child of wrapper.content) {
      if (child.type === "run") {
        runs.push(child);
        continue;
      }
      flush(renderedStart(child));
      rendered += renderChild(child);
    }
    flush(undefined);
    return rendered;
  };
  if (ctx.opts.trackedChanges === "clean") {
    // Insertions become real text; deletions vanish.
    if (wrapper.type === "insertion" || wrapper.type === "moveTo") {
      return renderChildren();
    }
    return "";
  }
  const inner = renderChildren();
  switch (wrapper.type) {
    case "insertion":
      return wrapInsertion(ctx, wrapper.info, inner);
    case "deletion":
      return wrapDeletion(ctx, wrapper.info, inner);
    case "moveFrom":
      return wrapMoveFrom(ctx, wrapper.info, inner);
    default:
      return wrapMoveTo(ctx, wrapper.info, inner);
  }
}

type CommentSlot = {
  start: number;
  comment?: Comment | undefined;
};

/**
 * Render the full inline content of a paragraph, tracking comment-range
 * boundaries to apply the configured wrapper.
 */
export function renderParagraphInline(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  content: readonly ParagraphContent[],
  paraId: string | undefined,
  edges: RunNeighbors = { before: "", after: "" },
): string {
  let out = "";
  // Stack of open comment ranges (document order) so nested comments wrap right.
  const openComments: CommentSlot[] = [];
  // Markers that write nothing are left out, so the runs on either side of
  // one meet as they do on the line.
  const items = content.filter(
    (item) =>
      !SILENT_MARKERS.has(item.type) &&
      !(ctx.opts.comments === "strip" && COMMENT_MARKERS.has(item.type)),
  );
  const around = (index: number): RunNeighbors => {
    const next = items[index + 1];
    return {
      before: out.length > 0 ? out.at(-1) : edges.before,
      after: next === undefined ? edges.after : renderedStart(next),
    };
  };

  for (let index = 0; index < items.length; index++) {
    const item = items[index];
    if (item === undefined) break;
    switch (item.type) {
      case "run": {
        // Consecutive runs are written together, so their emphasis joins up.
        const runs: Run[] = [item];
        while (items[index + 1]?.type === "run") {
          runs.push(items[index + 1] as Run);
          index++;
        }
        out += renderRuns(ctx, pkg, runs, paraId, around(index));
        break;
      }
      case "hyperlink":
        out += renderHyperlink(ctx, pkg, item, paraId);
        break;
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        out += renderTrackedWrapper(ctx, pkg, item, paraId);
        break;
      case "commentRangeStart":
        out += handleCommentStart(ctx, pkg, item, openComments, out.length);
        break;
      case "commentRangeEnd":
        out = handleCommentEnd(ctx, item, openComments, out);
        break;
      case "commentReference":
        // A point comment (or a range the parser collapsed to a reference).
        // It has no covered text, so emit just the marker.
        out += renderPointComment(ctx, pkg, item.id);
        break;
      case "simpleField":
      case "complexField": {
        // Render the visible result content.
        const runs = item.type === "simpleField" ? item.content : item.fieldResult;
        for (const child of runs) {
          if (child.type === "run") {
            out += renderRuns(ctx, pkg, [child], paraId);
            continue;
          }
          if (child.type === "hyperlink") {
            out += renderHyperlink(ctx, pkg, child, paraId);
            continue;
          }
          // A transparent wrapper carries no markdown of its own, so it is
          // read through to the content it holds; a capture contributes
          // whatever text it puts on the line.
          out +=
            child.type === "inlineWrapper"
              ? renderParagraphInline(ctx, pkg, child.content, paraId, UNKNOWN_NEIGHBORS)
              : child.text;
        }
        break;
      }
      // A content control states what its text is bound to and a bidirectional
      // wrapper states how it is laid out; markdown carries neither, so both
      // are read through to the text itself.
      case "inlineSdt":
      case "inlineWrapper":
        out += renderParagraphInline(ctx, pkg, item.content, paraId, around(index));
        break;
      case "mathEquation":
        // Markdown can't carry OMML; emit the plain-text fallback when present.
        if (item.plainText) {
          out += escapeInline(item.plainText);
        }
        break;
      default:
        // Range markers without inline payload (bookmark*, move*Range*,
        // commentReference, math) contribute nothing here.
        break;
    }
  }

  // Close any still-open comment ranges defensively.
  while (openComments.length) {
    const slot = openComments.pop();
    if (!slot) {
      break;
    }
    if (!slot.comment || ctx.opts.comments === "strip") {
      continue;
    }
    out = applyCommentWrapping(ctx, slot, out);
  }

  return out;
}

function handleCommentStart(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  marker: CommentRangeStart,
  openComments: CommentSlot[],
  startPos: number,
): string {
  if (ctx.opts.comments === "strip") {
    openComments.push({ start: startPos, comment: undefined });
    return "";
  }
  const comment = pkg?.document.comments?.find((c) => c.id === marker.id);
  openComments.push({ start: startPos, comment });
  return "";
}

function renderPointComment(ctx: RenderContext, pkg: DocxPackage | undefined, id: number): string {
  if (ctx.opts.comments === "strip") {
    return "";
  }
  const comment = pkg?.document.comments?.find((c) => c.id === id);
  if (!comment) {
    return "";
  }
  if (ctx.opts.comments === "sidecar") {
    const markerNumber = ctx.commentRefs.length + 1;
    ctx.commentRefs.push({ commentId: comment.id, markerNumber });
    return `[^c${markerNumber}]`;
  }
  // Inline: a point comment covers no text, so wrap an empty span.
  return wrapComment(ctx, { id: comment.id, author: comment.author }, "");
}

function handleCommentEnd(
  ctx: RenderContext,
  _marker: CommentRangeEnd,
  openComments: CommentSlot[],
  current: string,
): string {
  const slot = openComments.pop();
  if (!slot || !slot.comment || ctx.opts.comments === "strip") {
    return current;
  }
  return applyCommentWrapping(ctx, slot, current);
}

function applyCommentWrapping(ctx: RenderContext, slot: CommentSlot, current: string): string {
  if (!slot.comment) {
    return current;
  }
  const before = current.slice(0, slot.start);
  const inner = current.slice(slot.start);
  if (ctx.opts.comments === "sidecar") {
    const markerNumber = ctx.commentRefs.length + 1;
    ctx.commentRefs.push({ commentId: slot.comment.id, markerNumber });
    return `${before}${inner}[^c${markerNumber}]`;
  }
  return before + wrapComment(ctx, { id: slot.comment.id, author: slot.comment.author }, inner);
}
