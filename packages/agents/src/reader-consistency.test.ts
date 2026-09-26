/**
 * Cross-reader consistency: every surface that reads a document to a person or
 * a model must agree on each block's heading level and the number a reader
 * sees beside it. Each reader has its own tests, and each passed while two of
 * them disagreed about the same numbered heading (`## Scope` in the Markdown,
 * `1.` in `getContent()`), because no test read one document through more
 * than one of them.
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

/**
 * A synthetic numbered contract in the shapes real ones use:
 *
 * - `Heading 2` / `Heading 3` numbered through their styles' `w:numPr`
 *   (`1.`, `1.1.`), with a body paragraph continuing the same instance at the
 *   heading's child level, so a heading must advance the counters the list
 *   items after it read;
 * - a heading numbered by a direct `w:numPr`, and an outline-level paragraph
 *   (no heading style) numbered directly;
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
    ],
    nums: [
      ...(pkg.numbering?.nums ?? []),
      { numId: CLAUSES, abstractNumId: CLAUSES },
      { numId: ARTICLES, abstractNumId: ARTICLES },
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

  return (await ensureParaIds(new Uint8Array(await createDocx(document)))).docx;
};

/** What one reader says about one block. */
type BlockView = {
  text: string;
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
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
}): BlockView => {
  const number =
    block.displayLabel !== undefined && block.displayLabel !== block.styleId
      ? normalizeNumber(block.displayLabel)
      : undefined;
  return blockView({ text: block.text, headingLevel: block.headingLevel, number });
};

/**
 * `docxToMarkdown`, one non-blank line per block (the fixture has no
 * multi-line paragraphs), matched to the blocks in document order.
 */
const markdownViews = (markdown: string, texts: readonly string[]): BlockView[] => {
  const lines = markdown.split("\n").filter((line) => line.trim().length > 0);
  expect(lines).toHaveLength(texts.length);
  return lines.map((line, index) => {
    const text = texts[index] ?? "";
    expect(line.endsWith(text)).toBe(true);
    const prefix = line.slice(0, line.length - text.length).trim();
    const heading = /^(?<hashes>#{1,6})(?:\s+(?<marker>.*))?$/u.exec(prefix);
    if (heading?.groups) {
      return blockView({
        text,
        headingLevel: heading.groups["hashes"]?.length,
        number: normalizeNumber(heading.groups["marker"]),
      });
    }
    return blockView({
      text,
      headingLevel: undefined,
      number: normalizeNumber(prefix === "-" ? "•" : prefix),
    });
  });
};

type BlockViewOptions = {
  text: string;
  headingLevel: number | undefined;
  number: string | undefined;
};

/** A view without the fields the reader left undefined. */
const blockView = ({ text, headingLevel, number }: BlockViewOptions): BlockView => {
  const view: BlockView = { text };
  if (headingLevel !== undefined) {
    view.headingLevel = headingLevel;
  }
  if (number !== undefined) {
    view.number = number;
  }
  return view;
};

const readAll = async (bytes: Uint8Array) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(toArrayBuffer(bytes));
  const content = reviewer.getContent().filter(isFolioAIContentBlock);
  const snapshot = createReviewerBridge(reviewer).snapshot().blocks.filter(isFolioAIContentBlock);
  const markdown = await docxToMarkdown(toArrayBuffer(bytes), {
    annotations: "strip",
    trackedChanges: "clean",
    comments: "strip",
    footnotes: "keep",
  });
  const texts = content.map(({ text }) => text);
  return {
    getContent: content.map(snapshotView),
    snapshot: snapshot.map(snapshotView),
    markdown: markdownViews(markdown, texts),
  };
};

describe("readers agree on numbered headings and lists", () => {
  test("every reader shows each block's heading level and number alike", async () => {
    const views = await readAll(await buildNumberedDocument());

    // The fixture's intent, pinned once so a reader that drifts shows as a
    // disagreement below rather than rewriting the expectation.
    expect(views.getContent).toEqual([
      { text: "Agreement", headingLevel: 1 },
      { text: "Scope", headingLevel: 2, number: "1." },
      { text: "Definitions", headingLevel: 3, number: "1.1." },
      { text: "Terms", headingLevel: 3, number: "1.2." },
      { text: "Clause body item", number: "1.3." },
      { text: "The Supplier delivers the goods." },
      { text: "Direct article", headingLevel: 4, number: "(1)" },
      { text: "Outline article", headingLevel: 4, number: "a)" },
      { text: "a bullet", number: BULLET },
      { text: "another bullet", number: BULLET },
      { text: "first step", number: "1." },
      { text: "second step", number: "2." },
      { text: "Payment", headingLevel: 2, number: "2." },
      { text: "Payment is due in ten days (see clause 1)." },
    ]);
    expect(views.snapshot).toEqual(views.getContent);
    expect(views.markdown).toEqual(views.getContent);
  });
});
