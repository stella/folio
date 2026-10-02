/**
 * Cross-reader agreement over one saved package: `getContent()`, the AI
 * snapshot a bridge reads, the `read_document` rows a model reads, and
 * `docxToMarkdown`. The block view and its normalization (a bullet is "a
 * bullet" whatever its glyph; a heading's style-id label is not a number)
 * follow `packages/agents/src/reader-consistency.test.ts`, which pins one
 * numbered fixture from source; this one runs over whatever a scenario
 * produced, from the packed packages.
 */

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";
import { docxToMarkdown, isFolioAIContentBlock } from "@stll/folio-core/server";
import { XMLParser } from "fast-xml-parser";
import { marked, type Token } from "marked";

import { openReviewer, toArrayBuffer } from "./documents.ts";
import { featureIndex } from "./targets.ts";

/** What one reader says about one block. */
export type BlockView = {
  text: string;
  kind: "heading" | "listItem" | "paragraph";
  /** One-based heading level; absent when the reader shows no heading. */
  headingLevel?: number;
  /** The number or bullet a reader sees beside the text; absent when none. */
  number?: string;
};

type ContentBlock = {
  id: string;
  kind: string;
  text: string;
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
  listLevel?: number;
  table?: unknown;
  listReference?: { numId: number; level: number };
};

/** A bullet reads as `-` in Markdown and as its glyph elsewhere. */
export const BULLET = "(bullet)";

const normalizeNumber = (
  marker: string | undefined,
  format: string | undefined,
): string | undefined => {
  if (marker === undefined || marker.length === 0) {
    return undefined;
  }
  return format === "bullet" ? BULLET : marker;
};

const view = (
  text: string,
  kind: BlockView["kind"],
  headingLevel: number | undefined,
  number: string | undefined,
): BlockView => ({
  text,
  kind,
  ...(headingLevel === undefined ? {} : { headingLevel }),
  ...(number === undefined ? {} : { number }),
});

const asKind = (kind: string): BlockView["kind"] =>
  kind === "heading" || kind === "listItem" ? kind : "paragraph";

/** `getContent()` / snapshot block → view. Number formats come from package numbering. */
export const contentView = (
  block: ContentBlock,
  numberingFormats: ReadonlyMap<string, string>,
): BlockView => {
  const format =
    block.listReference === undefined
      ? undefined
      : numberingFormats.get(`${block.listReference.numId}:${block.listReference.level}`);
  if (block.listReference !== undefined && format === undefined) {
    throw new Error(
      `Missing numbering format for ${block.listReference.numId}:${block.listReference.level}`,
    );
  }
  const hasNumberLabel = block.displayLabel !== undefined && block.displayLabel !== block.styleId;
  if (hasNumberLabel && format === undefined) {
    throw new Error(`Missing numbering format for displayed label ${block.displayLabel}`);
  }
  return view(
    // Markdown keeps no leading or trailing blanks; neither side is compared on them.
    block.text.trim(),
    asKind(block.kind),
    block.headingLevel,
    hasNumberLabel ? normalizeNumber(block.displayLabel, format) : undefined,
  );
};

/** `[^1]: note text`, the note trailer after the body. */
const NOTE_REFERENCE = /\[\^[^\]]+\]/gu;

/**
 * Markdown writes a note reference as `[^n]` in document order; the block
 * text carries the reference mark itself. A line whose references stand where
 * the block text has a short mark reads as that block text.
 */
const matchNoteReferences = (line: string, text: string): string => {
  const [head = "", ...rest] = line.split(NOTE_REFERENCE);
  if (rest.length === 0) return line;
  const escape = (part: string) => part.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
  const mark = "[\\p{L}\\p{N}*†‡]{1,6}";
  const match = new RegExp(
    `^(?<start>.*)${rest.map((part) => mark + escape(part)).join("")}$`,
    "u",
  ).exec(text);
  const start = match?.groups?.["start"];
  return start !== undefined && head.endsWith(start)
    ? head.slice(0, head.length - start.length) + text
    : line;
};

const LIST_MARKER = /^(?<marker>\S+)\s+$/u;
const HTML_CELL_INLINE_TAGS = new Set([
  "a",
  "strong",
  "em",
  "u",
  "sup",
  "sub",
  "ins",
  "del",
  "span",
  "p",
]);

const inlineText = (tokens: readonly Token[]): string =>
  tokens
    .map((token) => {
      switch (token.type) {
        case "text":
          return token.tokens === undefined ? token.text : inlineText(token.tokens);
        case "escape":
        case "codespan":
          return token.text;
        case "strong":
        case "em":
        case "del":
        case "link":
          return inlineText(token.tokens);
        case "image":
          return "";
        case "br":
          return "\n";
        case "html": {
          const html = token.text;
          if (/^<\/?(?:u|sup|sub)>$/iu.test(html)) return "";
          throw new Error(`Unsupported inline Markdown HTML: ${html}`);
        }
        default:
          throw new Error(`Unsupported inline Markdown token: ${token.type}`);
      }
    })
    .join("");

const tokenText = (token: Token): string => {
  switch (token.type) {
    case "paragraph":
    case "heading":
      return inlineText(token.tokens);
    case "text":
      return inlineText(token.tokens === undefined ? [token] : token.tokens);
    case "code":
      return token.text;
    default:
      throw new Error(`Markdown token ${token.type} has no inline text`);
  }
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const htmlCellText = (nodes: readonly unknown[]): string => {
  const rendered: string[] = [];
  for (const node of nodes) {
    if (!isRecord(node)) throw new Error("Malformed HTML table node");
    for (const [tag, children] of Object.entries(node)) {
      if (tag === ":@") continue;
      if (tag === "#text") {
        if (typeof children !== "string") throw new Error("Malformed HTML table text");
        rendered.push(children);
        continue;
      }
      if (tag === "img") continue;
      if (tag === "br") {
        rendered.push("\n");
        continue;
      }
      if (!HTML_CELL_INLINE_TAGS.has(tag)) {
        throw new Error(`Unsupported HTML table cell tag: ${tag}`);
      }
      if (!Array.isArray(children)) throw new Error(`Malformed HTML table ${tag}`);
      rendered.push(htmlCellText(children));
    }
  }
  return rendered.join("");
};

const htmlTableViews = (html: string): BlockView[] => {
  const parser = new XMLParser({
    ignoreAttributes: false,
    preserveOrder: true,
    trimValues: false,
    unpairedTags: ["br", "img"],
  });
  const parsed: unknown = parser.parse(html);
  if (!Array.isArray(parsed)) throw new Error("Malformed HTML table");
  const views: BlockView[] = [];
  let sawTable = false;
  const walk = (nodes: readonly unknown[]): void => {
    for (const node of nodes) {
      if (!isRecord(node)) throw new Error("Malformed HTML table node");
      for (const [tag, children] of Object.entries(node)) {
        if (tag === ":@") continue;
        if (tag === "#text") {
          if (typeof children !== "string" || children.trim().length > 0) {
            throw new Error("Unexpected text outside an HTML table cell");
          }
          continue;
        }
        if (tag === "table" || tag === "thead" || tag === "tbody" || tag === "tr") {
          if (!Array.isArray(children)) throw new Error(`Malformed HTML table ${tag}`);
          if (tag === "table") sawTable = true;
          walk(children);
          continue;
        }
        if (tag === "th" || tag === "td") {
          if (!Array.isArray(children)) throw new Error(`Malformed HTML table ${tag}`);
          const text = htmlCellText(children).trim();
          if (text.length > 0) views.push(view(text, "paragraph", undefined, undefined));
          continue;
        }
        throw new Error(`Unsupported HTML table tag: ${tag}`);
      }
    }
  };
  walk(parsed);
  if (!sawTable) throw new Error("Markdown HTML block has no table");
  return views;
};

const suffixView = (
  rendered: string,
  expected: BlockView,
  kind: BlockView["kind"],
  headingLevel?: number,
): BlockView => {
  const plain = matchNoteReferences(rendered, expected.text);
  const suffix = expected.text.length > 0 && plain.endsWith(expected.text) ? expected.text : plain;
  const prefix = suffix === plain ? "" : plain.slice(0, plain.length - suffix.length);
  const marker = LIST_MARKER.exec(prefix)?.groups?.["marker"];
  if (kind === "paragraph" && expected.kind === "listItem" && marker !== undefined) {
    return view(suffix, "listItem", undefined, expected.number === BULLET ? BULLET : marker);
  }
  const headingMarker = kind === "heading" ? prefix.trim() : undefined;
  let number =
    headingMarker === undefined || headingMarker.length === 0 ? undefined : headingMarker;
  if (number !== undefined && expected.number === BULLET) number = BULLET;
  return view(suffix.trim(), kind, headingLevel, number);
};

/**
 * `docxToMarkdown` → views, matched against the texts the content reader
 * reports (a Markdown line carries no block boundary of its own). Pipe-table
 * rows yield one view per cell.
 */
export const markdownViews = (
  markdown: string,
  expectedViews: readonly BlockView[],
): BlockView[] => {
  const views: BlockView[] = [];
  const tokens = marked.lexer(markdown);
  let cursor = 0;
  const nextExpected = (): BlockView => {
    const expected = expectedViews.at(cursor);
    if (!expected) throw new Error("Markdown contains more text blocks than the source reader");
    cursor += 1;
    return expected;
  };
  const walk = (token: Token): void => {
    switch (token.type) {
      case "space":
      case "def":
        return;
      case "heading": {
        const expected = nextExpected();
        views.push(suffixView(tokenText(token), expected, "heading", token.depth));
        return;
      }
      case "paragraph": {
        const expected = nextExpected();
        views.push(suffixView(tokenText(token), expected, "paragraph"));
        return;
      }
      case "list": {
        for (const item of token.items) {
          for (const child of item.tokens) {
            if (child.type === "list") {
              walk(child);
              continue;
            }
            if (child.type !== "text" && child.type !== "paragraph") {
              throw new Error(`Unsupported list item Markdown token: ${child.type}`);
            }
            const expected = nextExpected();
            const text = tokenText(child);
            const marker = /^\s*(?<number>\d+[.)])\s+/u.exec(item.raw)?.groups?.["number"];
            if (token.ordered && marker === undefined) {
              throw new Error("Ordered Markdown list item has no rendered number");
            }
            const number = token.ordered ? marker : BULLET;
            const prefixLength = text.startsWith(expected.text) ? expected.text.length : 0;
            const firstRemainder = text.slice(prefixLength);
            const firstCustomMarker = /^\n(?<marker>[^\s]+?[.)])\s+/u.exec(firstRemainder);
            const firstNextExpected = expectedViews.at(cursor);
            const customContinuation =
              firstCustomMarker?.groups?.["marker"] !== undefined &&
              firstNextExpected?.kind === "listItem" &&
              (firstRemainder.slice(firstCustomMarker[0].length) === firstNextExpected.text ||
                firstRemainder
                  .slice(firstCustomMarker[0].length)
                  .startsWith(`${firstNextExpected.text}\n`));
            if (expected.text.length > 0 && text.startsWith(expected.text) && customContinuation) {
              views.push(view(expected.text, "listItem", undefined, number));
              let remainder = text.slice(expected.text.length);
              while (remainder.length > 0) {
                const customMarker = /^\n(?<marker>[^\s]+?[.)])\s+/u.exec(remainder);
                const nextSource = expectedViews.at(cursor);
                if (
                  customMarker?.groups?.["marker"] === undefined ||
                  nextSource?.kind !== "listItem"
                ) {
                  break;
                }
                const tail = remainder.slice(customMarker[0].length);
                if (tail !== nextSource.text && !tail.startsWith(`${nextSource.text}\n`)) break;
                cursor += 1;
                views.push(
                  suffixView(
                    `${customMarker.groups["marker"]} ${nextSource.text}`,
                    nextSource,
                    "paragraph",
                  ),
                );
                remainder = tail.slice(nextSource.text.length);
              }
              if (remainder.length > 0) {
                throw new Error("Unsupported text after a custom list marker");
              }
            } else {
              const itemView = suffixView(text, expected, "listItem");
              views.push(view(itemView.text, "listItem", undefined, number));
            }
          }
        }
        return;
      }
      case "table": {
        for (const cell of [...token.header, ...token.rows.flat()]) {
          const text = inlineText(cell.tokens);
          if (text.length === 0) continue;
          views.push(view(text.trim(), "paragraph", undefined, undefined));
          nextExpected();
        }
        return;
      }
      case "code": {
        const expected = nextExpected();
        views.push(view(token.text, "paragraph", undefined, undefined));
        if (expected.text.length === 0) throw new Error("Code block has no source paragraph");
        return;
      }
      case "hr":
      case "blockquote":
        throw new Error(`Unsupported structural Markdown token: ${token.type}`);
      case "html": {
        if (!token.block || !/<table(?:\s|>)/iu.test(token.text)) {
          throw new Error("Unsupported structural Markdown token: html");
        }
        for (const tableView of htmlTableViews(token.text)) {
          if (tableView.text.length > 0) nextExpected();
          views.push(tableView);
        }
        return;
      }
      default:
        throw new Error(`Unsupported block Markdown token: ${token.type}`);
    }
  };
  tokens.forEach(walk);
  if (cursor !== expectedViews.length) {
    throw new Error("Markdown contains fewer text blocks than the source reader");
  }
  return views;
};

/** Table cells are plain text in Markdown, whatever the cell paragraph is. */
const markdownComparable = (block: BlockView, inTable: boolean): BlockView =>
  inTable
    ? view(
        block.number === undefined ? block.text : `${block.number} ${block.text}`,
        "paragraph",
        undefined,
        undefined,
      )
    : block;

type LabelFields = { displayLabel?: string; headingLevel?: number; listLevel?: number };

/** The label fields a block or row carries, without the absent ones. */
export const labelFields = (source: Record<string, unknown>): LabelFields => {
  const fields: LabelFields = {};
  if (typeof source["displayLabel"] === "string") fields.displayLabel = source["displayLabel"];
  if (typeof source["headingLevel"] === "number") fields.headingLevel = source["headingLevel"];
  if (typeof source["listLevel"] === "number") fields.listLevel = source["listLevel"];
  return fields;
};

export type ReaderViews = {
  getContent: BlockView[];
  snapshot: BlockView[];
  readDocument: BlockView[];
  /** `getContent()` of the package with every change accepted, as Markdown can express it. */
  getContentAsMarkdown: BlockView[];
  markdown: BlockView[];
  /** `read_document` rows, raw, for the fields a view does not carry. */
  rows: Record<string, unknown>[];
  /** `getContent()` ids, in order. */
  ids: string[];
  /** `getContent()`'s own label fields, which `read_document` rows restate. */
  labels: LabelFields[];
};

export const MARKDOWN_READ_OPTIONS = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  footnotes: "keep",
} as const;

/** Read one package through every reader. */
export const readAll = async (bytes: Uint8Array): Promise<ReaderViews> => {
  const reviewer = await openReviewer(bytes);
  const numberingFormats = new Map(
    reviewer
      .readNumberingDefinitions()
      .map(({ numId, level, format }) => [`${numId}:${level}`, format] as const),
  );
  const content = (reviewer.getContent() as ContentBlock[]).filter((block) =>
    isFolioAIContentBlock(block as never),
  );
  const bridge = createReviewerBridge(reviewer);
  const snapshot = (bridge.snapshot().blocks as ContentBlock[]).filter((block) =>
    isFolioAIContentBlock(block as never),
  );
  const read = executeFolioToolCallUntyped("read_document", {}, bridge, {});
  if (!read.ok) {
    throw new Error(`read_document failed: ${read.error}`);
  }
  const rows = read.result as Record<string, unknown>[];
  const markdown = await docxToMarkdown(toArrayBuffer(bytes), MARKDOWN_READ_OPTIONS);
  const contentViews = content.map((block) => contentView(block, numberingFormats));
  // Clean Markdown shows every change accepted, structure included (a
  // deleted paragraph mark joins two paragraphs); the readers above show the
  // markup. Markdown is compared with the same package, accepted.
  let accepted = content;
  let acceptedReviewer = reviewer;
  if (reviewer.getChanges().length > 0) {
    acceptedReviewer = await openReviewer(bytes);
    acceptedReviewer.acceptAll();
    accepted = (acceptedReviewer.getContent() as ContentBlock[]).filter((block) =>
      isFolioAIContentBlock(block as never),
    );
  }
  // `docxToMarkdown` writes no text-box paragraph the block readers list
  // (MARKDOWN_DROPS_TEXT_BOX, pinned in known-issues.test.ts); the rest of
  // the document is still compared. Drop this once that finding is fixed.
  const boxed = featureIndex(acceptedReviewer).inTextBox;
  accepted = accepted.filter((block) => !boxed.has(block.id));
  return {
    getContent: contentViews,
    snapshot: snapshot.map((block) => contentView(block, numberingFormats)),
    readDocument: rows.map((row) =>
      view(String(row["text"]).trim(), asKind(String(row["kind"])), undefined, undefined),
    ),
    getContentAsMarkdown: accepted.map((block) =>
      markdownComparable(contentView(block, numberingFormats), block.table !== undefined),
    ),
    markdown: markdownViews(
      markdown,
      accepted.map((block) => contentView(block, numberingFormats)),
    ),
    rows,
    ids: content.map(({ id }) => id),
    labels: content.map(labelFields),
  };
};
