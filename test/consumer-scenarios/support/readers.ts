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
};

/** A bullet reads as `-` in Markdown and as its glyph elsewhere. */
export const BULLET = "(bullet)";

const normalizeNumber = (marker: string | undefined): string | undefined => {
  if (marker === undefined || marker.length === 0) {
    return undefined;
  }
  return /[\p{L}\p{N}]/u.test(marker) ? marker : BULLET;
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

/** `getContent()` / snapshot block → view. A style-id label is not a number. */
export const contentView = (block: ContentBlock): BlockView =>
  view(
    // Markdown keeps no leading or trailing blanks; neither side is compared on them.
    block.text.trim(),
    asKind(block.kind),
    block.headingLevel,
    block.displayLabel !== undefined && block.displayLabel !== block.styleId
      ? normalizeNumber(block.displayLabel)
      : undefined,
  );

const MARKDOWN_ESCAPE = /\\(?<char>[\\`*_{}[\]()#+\-.!|<>~])/gu;
/** Inline emphasis a formatting edit adds; the block text carries none of it. */
const EMPHASIS =
  /\*\*|__|~~|<\/?(?:u|sup|sub)>|(?<![\p{L}\p{N}\\])[*_]|(?<!\\)[*_](?![\p{L}\p{N}])/gu;
/**
 * A hyperlink reads as its text; the block text carries no target. An empty
 * one (a split at its edge leaves one) reads as nothing, so the blanks
 * before it are trailing blanks again.
 */
const LINK = /(?<![!\\])\[(?<label>(?:[^\]\\]|\\.)*)\]\((?:[^()\s\\]|\\.|\([^()\s]*\))*\)/gu;
const unescapeMarkdown = (text: string): string => {
  const unlinked = text.replace(LINK, "$<label>");
  return (unlinked === text ? text : unlinked.trim())
    .replace(EMPHASIS, "")
    .replace(MARKDOWN_ESCAPE, "$<char>");
};

/** `[^1]: note text`, the note trailer after the body. */
const NOTE_DEFINITION = /^\[\^[^\]]+\]:/u;
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

const LIST_MARKER = /^(?<indent>\s*)(?<marker>[-*+]|\S+?[.)]|\(\S+?\)|\S)\s+(?<rest>.*)$/u;

/**
 * `docxToMarkdown` → views, matched against the texts the content reader
 * reports (a Markdown line carries no block boundary of its own). Pipe-table
 * rows yield one view per cell.
 */
export const markdownViews = (markdown: string, texts: readonly string[]): BlockView[] => {
  const views: BlockView[] = [];
  let cursor = 0;
  const nextText = (): string => texts[cursor++] ?? "";
  for (const rawLine of markdown.split("\n")) {
    const line = rawLine.trimEnd();
    if (line.trim().length === 0) continue;
    if (line.trimStart().startsWith("|")) {
      const cells = line
        .trim()
        .slice(1, -1)
        .split(/(?<!\\)\|/u);
      if (cells.every((cell) => /^\s*:?-{3,}:?\s*$/u.test(cell))) continue;
      for (const cell of cells) {
        const text = unescapeMarkdown(cell.trim());
        // A blank cell paragraph is structure, not a content block.
        if (text.length === 0) continue;
        views.push(view(text, "paragraph", undefined, undefined));
        nextText();
      }
      continue;
    }
    if (NOTE_DEFINITION.test(line)) continue;
    // An empty list item or heading is a blank paragraph, which the content
    // readers leave out as structure.
    if (
      /^\s*(?:[-*+]|#{1,6}|\(?(?:\d{1,4}(?:\.\d{1,4})*|[a-zA-Z]|[ivxlcdmIVXLCDM]{1,6})[.)])\s*$/u.test(
        line,
      )
    ) {
      continue;
    }
    const text = nextText();
    const plain = matchNoteReferences(unescapeMarkdown(line), text);
    const prefix = plain.endsWith(text) ? plain.slice(0, plain.length - text.length).trim() : null;
    if (prefix === null) {
      views.push(view(plain, "paragraph", undefined, undefined));
      continue;
    }
    const heading = /^(?<hashes>#{1,6})(?:\s+(?<marker>.*))?$/u.exec(prefix);
    if (heading?.groups) {
      views.push(
        view(
          text,
          "heading",
          heading.groups["hashes"]?.length,
          normalizeNumber(heading.groups["marker"]),
        ),
      );
      continue;
    }
    if (prefix.length === 0) {
      views.push(view(text, "paragraph", undefined, undefined));
      continue;
    }
    const marker = LIST_MARKER.exec(`${prefix} ${text}`)?.groups?.["marker"] ?? prefix;
    views.push(view(text, "listItem", undefined, normalizeNumber(marker === "-" ? "•" : marker)));
  }
  return views;
};

/** Table cells are plain text in Markdown, whatever the cell paragraph is. */
const markdownComparable = (block: BlockView, inTable: boolean): BlockView =>
  inTable ? view(block.text, "paragraph", undefined, undefined) : block;

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
  const contentViews = content.map(contentView);
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
    snapshot: snapshot.map(contentView),
    readDocument: rows.map((row) =>
      view(String(row["text"]).trim(), asKind(String(row["kind"])), undefined, undefined),
    ),
    getContentAsMarkdown: accepted.map((block) =>
      markdownComparable(contentView(block), block.table !== undefined),
    ),
    markdown: markdownViews(
      markdown,
      accepted.map(({ text }) => text.trim()),
    ),
    rows,
    ids: content.map(({ id }) => id),
    labels: content.map(labelFields),
  };
};
