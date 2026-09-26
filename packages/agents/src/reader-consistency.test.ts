/**
 * Cross-reader consistency: every surface that reads a document to a person or
 * a model must agree on each block's kind, heading level and the number a
 * reader sees beside it. Each reader has its own tests, and each passed while
 * two of them disagreed about the same numbered heading (`## Scope` in the
 * Markdown, `1.` in `getContent()`; `listItem` in the snapshot, a heading
 * everywhere else), because no test read one document through more than one
 * of them.
 *
 * A reader added later (or a field a reader starts to expose) belongs in
 * `readAll` below; a numbering shape a reader got wrong belongs in
 * `buildNumberedDocument`.
 */

import { describe, expect, test } from "bun:test";

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createDocx,
  docxToMarkdown,
  ensureParaIds,
  FolioDocxReviewer,
  isFolioAIContentBlock,
} from "@stll/folio-core/server";

import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCall } from "./execute";
import type { FolioAgentBlock } from "./types";

type Document = ReturnType<typeof fromMarkdown>;
type Paragraph = Extract<Document["package"]["document"]["content"][number], { type: "paragraph" }>;
type NumberingLevel = NonNullable<
  NonNullable<Document["package"]["numbering"]>["abstractNums"]
>[number]["levels"][number];

/** A decimal level unless the options say otherwise. */
type LevelOptions = Pick<NumberingLevel, "ilvl" | "lvlText"> & Partial<NumberingLevel>;

const level = (options: LevelOptions): NumberingLevel => ({
  start: 1,
  numFmt: "decimal",
  suffix: "space",
  pPr: { indentLeft: 0, indentFirstLine: 0 },
  ...options,
});

/** The ids the fixture's numbering definitions use. */
const CLAUSES = 5;
const ARTICLES = 6;
const HIDDEN = 7;

/**
 * A synthetic numbered contract in the shapes real ones use:
 *
 * - `Heading 2` / `Heading 3` numbered through their styles' `w:numPr`
 *   (`1.`, `1.1.`), with a body paragraph continuing the same instance at the
 *   heading's child level, so a heading must advance the counters the list
 *   items after it read;
 * - a heading numbered by a direct `w:numPr`, and an outline-level paragraph
 *   (no heading style) numbered directly;
 * - a heading and a body paragraph on a level whose marker is hidden
 *   (`w:vanish`), and a `Heading 2` and a body paragraph that cancel
 *   numbering with the reserved `w:numId="0"`: numbered in the package, no
 *   number on the page;
 * - an ordinary bulleted and an ordinary numbered list.
 */
const buildNumberedDocument = async (): Promise<Uint8Array> => {
  const document = fromMarkdown(
    [
      "# Agreement",
      "## Scope",
      "### Definitions",
      "### Terms",
      "Clause body item",
      "The Supplier delivers the goods.",
      "Direct article",
      "Outline article",
      "Hidden-number heading",
      "Hidden-number paragraph",
      "Under a hidden number",
      "Unnumbered clause",
      "Cancelled item",
      "- a bullet\n- another bullet",
      "1. first step\n2. second step",
      "## Payment",
      "Payment is due in ten days (see clause 1).",
    ].join("\n\n"),
  );
  const pkg = document.package;
  pkg.numbering = {
    abstractNums: [
      ...(pkg.numbering?.abstractNums ?? []),
      {
        abstractNumId: CLAUSES,
        multiLevelType: "multilevel",
        levels: [level({ ilvl: 0, lvlText: "%1." }), level({ ilvl: 1, lvlText: "%1.%2." })],
      },
      {
        abstractNumId: ARTICLES,
        multiLevelType: "multilevel",
        levels: [
          level({ ilvl: 0, lvlText: "(%1)" }),
          level({ ilvl: 1, lvlText: "%2)", numFmt: "lowerLetter" }),
        ],
      },
      {
        abstractNumId: HIDDEN,
        multiLevelType: "multilevel",
        levels: [
          level({ ilvl: 0, lvlText: "Art. %1", rPr: { hidden: true } }),
          level({ ilvl: 1, lvlText: "%1.%2." }),
        ],
      },
    ],
    nums: [
      ...(pkg.numbering?.nums ?? []),
      { numId: CLAUSES, abstractNumId: CLAUSES },
      { numId: ARTICLES, abstractNumId: ARTICLES },
      { numId: HIDDEN, abstractNumId: HIDDEN },
    ],
  };

  const style = (styleId: string) => {
    const found = pkg.styles?.styles.find((candidate) => candidate.styleId === styleId);
    if (!found) {
      throw new Error(`fixture style ${styleId} is missing`);
    }
    return found;
  };
  const heading2 = style("Heading2");
  heading2.pPr = {
    ...heading2.pPr,
    numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 0 }),
  };
  const heading3 = style("Heading3");
  heading3.pPr = {
    ...heading3.pPr,
    numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 1 }),
  };

  const paragraph = (text: string): Paragraph => {
    const found = pkg.document.content.find(
      (block): block is Paragraph =>
        block.type === "paragraph" &&
        block.content.some(
          (item) =>
            item.type === "run" &&
            item.content.some((content) => content.type === "text" && content.text === text),
        ),
    );
    if (!found) {
      throw new Error(`fixture paragraph "${text}" is missing`);
    }
    return found;
  };
  const format = (text: string, formatting: NonNullable<Paragraph["formatting"]>): void => {
    const target = paragraph(text);
    target.formatting = { ...target.formatting, ...formatting };
  };
  format("Clause body item", { numPr: paragraphNumberingFromSlots({ numId: CLAUSES, ilvl: 1 }) });
  format("Direct article", {
    styleId: "Heading4",
    numPr: paragraphNumberingFromSlots({ numId: ARTICLES, ilvl: 0 }),
  });
  format("Outline article", {
    outlineLevel: { kind: "heading", level: 3 },
    numPr: paragraphNumberingFromSlots({ numId: ARTICLES, ilvl: 1 }),
  });
  format("Hidden-number heading", {
    styleId: "Heading4",
    numPr: paragraphNumberingFromSlots({ numId: HIDDEN, ilvl: 0 }),
  });
  format("Hidden-number paragraph", {
    numPr: paragraphNumberingFromSlots({ numId: HIDDEN, ilvl: 0 }),
  });
  format("Under a hidden number", {
    numPr: paragraphNumberingFromSlots({ numId: HIDDEN, ilvl: 1 }),
  });
  format("Unnumbered clause", {
    styleId: "Heading2",
    numPr: paragraphNumberingFromSlots({ numId: 0 }),
  });
  format("Cancelled item", { numPr: paragraphNumberingFromSlots({ numId: 0 }) });

  return (await ensureParaIds(new Uint8Array(await createDocx(document)))).docx;
};

type Kind = "heading" | "listItem" | "paragraph";

/** What one reader says about one block. */
type BlockView = {
  text: string;
  kind: Kind;
  /** One-based heading level; absent when the reader shows no heading. */
  headingLevel?: number;
  /** The number or bullet a reader sees beside the text; absent when none. */
  number?: string;
};

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

/** A bullet reads as `-` in Markdown and as its glyph elsewhere; both are "a bullet". */
const BULLET = "(bullet)";
const normalizeNumber = (marker: string | undefined): string | undefined => {
  if (marker === undefined || marker.length === 0) {
    return undefined;
  }
  return /[\p{L}\p{N}]/u.test(marker) ? marker : BULLET;
};

/**
 * `getContent()` / the AI snapshot. A heading's `displayLabel` is its style
 * id when it carries no number, so only a label that is not the style id is a
 * number a reader sees.
 */
const snapshotView = (block: {
  text: string;
  kind: string;
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
}): BlockView => {
  const number =
    block.displayLabel !== undefined && block.displayLabel !== block.styleId
      ? normalizeNumber(block.displayLabel)
      : undefined;
  return blockView({
    text: block.text,
    kind: asKind(block.kind),
    headingLevel: block.headingLevel,
    number,
  });
};

const asKind = (kind: string): Kind => {
  if (kind === "heading" || kind === "listItem" || kind === "paragraph") {
    return kind;
  }
  throw new Error(`unexpected block kind ${kind}`);
};

/** A line with no heading tag reads as a list item when it shows a number. */
const bodyView = (text: string, number: string | undefined): BlockView =>
  blockView({
    text,
    kind: number === undefined ? "paragraph" : "listItem",
    headingLevel: undefined,
    number,
  });

/**
 * `docxToMarkdown`, one non-blank line per block (the fixture has no
 * multi-line paragraphs), matched to the blocks in document order.
 */
const markdownViews = (markdown: string, texts: readonly string[]): BlockView[] =>
  linesFor({
    output: markdown,
    texts,
    view: (prefix, text) => {
      const heading = /^(?<hashes>#{1,6})(?:\s+(?<marker>.*))?$/u.exec(prefix);
      if (heading?.groups) {
        return blockView({
          text,
          kind: "heading",
          headingLevel: heading.groups["hashes"]?.length,
          number: normalizeNumber(heading.groups["marker"]),
        });
      }
      return bodyView(text, normalizeNumber(prefix === "-" ? "•" : prefix));
    },
  });

/** `getContentAsText()`: `[id] (h2) 1. Scope`, `[id] a) item`, `[id] text`. */
const contentTextViews = (lines: string, texts: readonly string[]): BlockView[] =>
  linesFor({
    output: lines,
    texts,
    view: (prefix, text) => {
      const line = /^\[[^\]]+\](?:\s+\(h(?<level>\d)\))?(?:\s+(?<marker>.*))?$/u.exec(prefix);
      if (!line?.groups) {
        throw new Error(`unexpected getContentAsText line prefix "${prefix}"`);
      }
      const number = normalizeNumber(line.groups["marker"]);
      const headingLevel = line.groups["level"];
      if (headingLevel !== undefined) {
        return blockView({ text, kind: "heading", headingLevel: Number(headingLevel), number });
      }
      return bodyView(text, number);
    },
  });

/** One view per non-blank output line, each line ending in its block's text. */
type LinesForOptions = {
  output: string;
  texts: readonly string[];
  view: (prefix: string, text: string) => BlockView;
};

const linesFor = ({ output, texts, view }: LinesForOptions): BlockView[] => {
  const lines = output.split("\n").filter((line) => line.trim().length > 0);
  expect(lines).toHaveLength(texts.length);
  return lines.map((line, index) => {
    const text = texts[index] ?? "";
    expect(line.endsWith(text)).toBe(true);
    return view(line.slice(0, line.length - text.length).trim(), text);
  });
};

type BlockViewOptions = {
  text: string;
  kind: Kind;
  headingLevel: number | undefined;
  number: string | undefined;
};

/** A view without the fields the reader left undefined. */
const blockView = ({ text, kind, headingLevel, number }: BlockViewOptions): BlockView => {
  const view: BlockView = { text, kind };
  if (headingLevel !== undefined) {
    view.headingLevel = headingLevel;
  }
  if (number !== undefined) {
    view.number = number;
  }
  return view;
};

/** The `read_document` row fields that restate a snapshot block's. */
const rowFields = (row: FolioAgentBlock) => ({
  blockId: row.blockId,
  kind: row.kind,
  displayLabel: row.displayLabel,
  headingLevel: row.headingLevel,
  listLevel: row.listLevel,
});

const readAll = async (bytes: Uint8Array) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(toArrayBuffer(bytes));
  const bridge = createReviewerBridge(reviewer);
  const content = reviewer.getContent().filter(isFolioAIContentBlock);
  const snapshot = bridge.snapshot().blocks.filter(isFolioAIContentBlock);
  const read = executeFolioToolCall("read_document", {}, bridge);
  if (!read.ok) {
    throw new Error(read.error);
  }
  const markdown = await docxToMarkdown(toArrayBuffer(bytes), {
    annotations: "strip",
    trackedChanges: "clean",
    comments: "strip",
    footnotes: "keep",
  });
  const texts = content.map(({ text }) => text);
  const styleIds = new Map(content.map((block) => [block.id, block.styleId]));
  return {
    content,
    rows: read.result,
    views: {
      getContent: content.map(snapshotView),
      snapshot: snapshot.map(snapshotView),
      readDocument: read.result.map((row) =>
        snapshotView({ ...row, styleId: styleIds.get(row.blockId) }),
      ),
      getContentAsText: contentTextViews(reviewer.getContentAsText(), texts),
      markdown: markdownViews(markdown, texts),
    },
  };
};

describe("readers agree on numbered headings and lists", () => {
  test("every reader shows each block's kind, heading level and number alike", async () => {
    const { views } = await readAll(await buildNumberedDocument());

    // The fixture's intent, pinned once so a reader that drifts shows as a
    // disagreement below rather than rewriting the expectation.
    expect(views.getContent).toEqual([
      { text: "Agreement", kind: "heading", headingLevel: 1 },
      { text: "Scope", kind: "heading", headingLevel: 2, number: "1." },
      { text: "Definitions", kind: "heading", headingLevel: 3, number: "1.1." },
      { text: "Terms", kind: "heading", headingLevel: 3, number: "1.2." },
      { text: "Clause body item", kind: "listItem", number: "1.3." },
      { text: "The Supplier delivers the goods.", kind: "paragraph" },
      { text: "Direct article", kind: "heading", headingLevel: 4, number: "(1)" },
      { text: "Outline article", kind: "heading", headingLevel: 4, number: "a)" },
      { text: "Hidden-number heading", kind: "heading", headingLevel: 4 },
      { text: "Hidden-number paragraph", kind: "paragraph" },
      // A hidden marker still counts: the child shows the parent's number.
      { text: "Under a hidden number", kind: "listItem", number: "2.1." },
      { text: "Unnumbered clause", kind: "heading", headingLevel: 2 },
      { text: "Cancelled item", kind: "paragraph" },
      { text: "a bullet", kind: "listItem", number: BULLET },
      { text: "another bullet", kind: "listItem", number: BULLET },
      { text: "first step", kind: "listItem", number: "1." },
      { text: "second step", kind: "listItem", number: "2." },
      { text: "Payment", kind: "heading", headingLevel: 2, number: "2." },
      { text: "Payment is due in ten days (see clause 1).", kind: "paragraph" },
    ]);
    for (const [reader, view] of Object.entries(views)) {
      expect({ reader, view }).toEqual({ reader, view: views.getContent });
    }
  });

  test("read_document rows restate the snapshot's label and levels", async () => {
    const { content, rows } = await readAll(await buildNumberedDocument());

    expect(rows.map(rowFields)).toEqual(
      content.map((block) => ({
        blockId: block.id,
        kind: block.kind,
        displayLabel: block.displayLabel,
        headingLevel: block.headingLevel,
        listLevel: block.listLevel,
      })),
    );
    // Rows stay compact: a field the block lacks is absent, not `undefined`.
    const body = rows.find(({ text }) => text === "The Supplier delivers the goods.");
    expect(body && Object.keys(body).sort()).toEqual(["blockId", "blockTextHash", "kind", "text"]);
    expect(rows.find(({ text }) => text === "Definitions")).toMatchObject({
      kind: "heading",
      displayLabel: "1.1.",
      headingLevel: 3,
      listLevel: 1,
    });
  });
});
