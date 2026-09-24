import { describe, expect, test } from "bun:test";

import { layoutDocument } from "./index";
import { calculateChainHeight, computeKeepNextChains } from "./keep-together";
import type { FlowBlock, Measure, ParagraphBlock, ParagraphMeasure } from "./types";

const paragraph = (id: string, keepNext = false): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text: id }],
  attrs: { keepNext },
});

const emptyParagraph = (id: string): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [],
});

const authoredEmptyParagraph = (id: string): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [],
  attrs: { hasDirectParagraphFormatting: true },
});

const table = (id: string): FlowBlock => ({
  kind: "table",
  id,
  rows: [],
});

const paragraphMeasure = (...lineHeights: number[]): ParagraphMeasure => ({
  kind: "paragraph",
  lines: lineHeights.map((lineHeight) => ({
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: 1,
    width: 10,
    ascent: lineHeight * 0.8,
    descent: lineHeight * 0.2,
    lineHeight,
  })),
  totalHeight: lineHeights.reduce((total, lineHeight) => total + lineHeight, 0),
});

type Spacing = { before: number; after: number };

const spacedParagraph = (id: string, spacing: Spacing, keepNext = false): ParagraphBlock => ({
  kind: "paragraph",
  id,
  pmStart: 0,
  pmEnd: 0,
  runs: [{ kind: "text", text: id }],
  attrs: { keepNext, spacing },
});

/** Measure shaped like the real measurer: totalHeight includes the paragraph's own spacing. */
const spacedParagraphMeasure = (spacing: Spacing, ...lineHeights: number[]): ParagraphMeasure => {
  const measure = paragraphMeasure(...lineHeights);
  return { ...measure, totalHeight: measure.totalHeight + spacing.before + spacing.after };
};

const pageGeometry = {
  pageSize: { w: 600, h: 1000 },
  margins: { top: 50, right: 50, bottom: 50, left: 50 },
};
const pageOptions = {
  ...pageGeometry,
  finalPageSize: pageGeometry.pageSize,
  finalMargins: pageGeometry.margins,
};

const pageBlockIds = (result: ReturnType<typeof layoutDocument>): string[][] =>
  result.pages.map((page) => page.fragments.map((fragment) => String(fragment.blockId)));

describe("calculateChainHeight", () => {
  test("counts each paragraph's spacing once when measures include it", () => {
    const headingSpacing = { before: 24, after: 8 };
    const anchorSpacing = { before: 8, after: 8 };
    const blocks: FlowBlock[] = [
      spacedParagraph("heading", headingSpacing, true),
      spacedParagraph("anchor", anchorSpacing),
    ];
    const measures: Measure[] = [
      spacedParagraphMeasure(headingSpacing, 24),
      spacedParagraphMeasure(anchorSpacing, 16, 16),
    ];
    const chain = { startIndex: 0, endIndex: 0, memberIndices: [0], anchorIndex: 1 };

    // before 24 + heading line 24 + collapsed gap max(8, 8) + both anchor lines,
    // since widow control cannot split a two-line anchor
    expect(calculateChainHeight(chain, blocks, measures)).toBe(88);
  });

  test("counts a fully reserved member's spacing once", () => {
    const spacing = { before: 12, after: 6 };
    const blocks: FlowBlock[] = [
      spacedParagraph("heading", spacing, true),
      spacedParagraph("subheading", spacing, true),
      spacedParagraph("anchor", spacing),
    ];
    const measures: Measure[] = [
      spacedParagraphMeasure(spacing, 20),
      spacedParagraphMeasure(spacing, 14),
      spacedParagraphMeasure(spacing, 10, 10),
    ];
    const chain = { startIndex: 0, endIndex: 1, memberIndices: [0, 1], anchorIndex: 2 };

    // 12 + 20 + max(12, 6) + 14 + max(12, 6) + 10 + 10
    expect(calculateChainHeight(chain, blocks, measures)).toBe(90);
  });

  test("collapses the first member's spacing before with incoming spacing after", () => {
    const blocks: FlowBlock[] = [
      spacedParagraph("heading", { before: 6, after: 0 }, true),
      spacedParagraph("anchor", { before: 0, after: 0 }),
    ];
    const measures: Measure[] = [
      spacedParagraphMeasure({ before: 6, after: 0 }, 20),
      spacedParagraphMeasure({ before: 0, after: 0 }, 10),
    ];
    const chain = { startIndex: 0, endIndex: 0, memberIndices: [0], anchorIndex: 1 };

    expect(calculateChainHeight(chain, blocks, measures, 18)).toBe(48);
  });

  test("reserves only the widow-controlled opening of a splittable successor", () => {
    const blocks: FlowBlock[] = [
      paragraph("first", true),
      paragraph("splittable", true),
      paragraph("anchor"),
    ];
    const measures: Measure[] = [
      paragraphMeasure(12, 12),
      paragraphMeasure(14, 14, 14, 14),
      paragraphMeasure(16),
    ];

    expect(
      calculateChainHeight(
        {
          startIndex: 0,
          endIndex: 1,
          memberIndices: [0, 1],
          anchorIndex: 2,
        },
        blocks,
        measures,
      ),
    ).toBe(52);
  });

  test("reserves every consecutive single-line member through the anchor", () => {
    const blocks: FlowBlock[] = [
      paragraph("empty heading", true),
      paragraph("heading", true),
      paragraph("anchor"),
    ];
    const measures: Measure[] = [paragraphMeasure(12), paragraphMeasure(14), paragraphMeasure(16)];

    expect(
      calculateChainHeight(
        {
          startIndex: 0,
          endIndex: 1,
          memberIndices: [0, 1],
          anchorIndex: 2,
        },
        blocks,
        measures,
      ),
    ).toBe(42);
  });

  test("reserves two anchor lines after a trailing table separator", () => {
    const blocks: FlowBlock[] = [table("table"), emptyParagraph("separator"), paragraph("body")];
    const measures: Measure[] = [
      { kind: "table", rows: [], columnWidths: [], totalWidth: 0, totalHeight: 100 },
      paragraphMeasure(12),
      paragraphMeasure(14, 14, 14, 14),
    ];

    expect(
      calculateChainHeight(
        {
          startIndex: 1,
          endIndex: 1,
          memberIndices: [1],
          anchorIndex: 2,
        },
        blocks,
        measures,
      ),
    ).toBe(40);
  });

  test("reserves one anchor line when widow control is disabled", () => {
    const anchor = paragraph("body");
    anchor.attrs = { widowControl: false };
    const blocks: FlowBlock[] = [table("table"), emptyParagraph("separator"), anchor];
    const measures: Measure[] = [
      { kind: "table", rows: [], columnWidths: [], totalWidth: 0, totalHeight: 100 },
      paragraphMeasure(12),
      paragraphMeasure(14, 14, 14),
    ];

    expect(
      calculateChainHeight(
        {
          startIndex: 1,
          endIndex: 1,
          memberIndices: [1],
          anchorIndex: 2,
        },
        blocks,
        measures,
      ),
    ).toBe(26);
  });

  test("reserves a whole widow-controlled anchor shorter than four lines", () => {
    const blocks: FlowBlock[] = [paragraph("heading", true), paragraph("body")];
    const measures: Measure[] = [paragraphMeasure(20), paragraphMeasure(10, 11, 12)];
    const chain = { startIndex: 0, endIndex: 0, memberIndices: [0], anchorIndex: 1 };

    expect(calculateChainHeight(chain, blocks, measures)).toBe(53);
  });

  test("reserves two lines of a widow-controlled anchor with four or more lines", () => {
    const blocks: FlowBlock[] = [paragraph("heading", true), paragraph("body")];
    const measures: Measure[] = [paragraphMeasure(20), paragraphMeasure(10, 11, 12, 13)];
    const chain = { startIndex: 0, endIndex: 0, memberIndices: [0], anchorIndex: 1 };

    expect(calculateChainHeight(chain, blocks, measures)).toBe(41);
  });

  test("reserves a whole keepLines anchor without widow control", () => {
    const anchor = paragraph("body");
    anchor.attrs = { keepLines: true, widowControl: false };
    const blocks: FlowBlock[] = [paragraph("heading", true), anchor];
    const measures: Measure[] = [paragraphMeasure(20), paragraphMeasure(10, 10, 10, 10, 10)];
    const chain = { startIndex: 0, endIndex: 0, memberIndices: [0], anchorIndex: 1 };

    expect(calculateChainHeight(chain, blocks, measures)).toBe(70);
  });

  test("continues through a widow-controlled member too short to split", () => {
    const blocks: FlowBlock[] = [
      paragraph("heading", true),
      paragraph("lead-in", true),
      paragraph("body"),
    ];
    const measures: Measure[] = [
      paragraphMeasure(20),
      paragraphMeasure(10, 10, 10),
      paragraphMeasure(5, 5, 5, 5, 5),
    ];
    const chain = { startIndex: 0, endIndex: 1, memberIndices: [0, 1], anchorIndex: 2 };

    expect(calculateChainHeight(chain, blocks, measures)).toBe(60);
  });
});

describe("computeKeepNextChains", () => {
  test("carries keepNext through an empty separator to the next content paragraph", () => {
    const blocks: FlowBlock[] = [
      paragraph("heading", true),
      emptyParagraph("separator"),
      paragraph("body"),
    ];

    expect(computeKeepNextChains(blocks).get(0)).toEqual({
      startIndex: 0,
      endIndex: 1,
      memberIndices: [0, 1],
      anchorIndex: 2,
    });
  });

  test("stops keepNext at an authored empty spacer", () => {
    const blocks: FlowBlock[] = [
      paragraph("heading", true),
      authoredEmptyParagraph("answer space"),
      paragraph("next heading", true),
      paragraph("body"),
    ];

    expect(computeKeepNextChains(blocks)).toEqual(
      new Map([
        [
          0,
          {
            startIndex: 0,
            endIndex: 0,
            memberIndices: [0],
            anchorIndex: 1,
          },
        ],
        [
          2,
          {
            startIndex: 2,
            endIndex: 2,
            memberIndices: [2],
            anchorIndex: 3,
          },
        ],
      ]),
    );
  });

  test("carries a trailing table separator to the next content paragraph", () => {
    const blocks: FlowBlock[] = [table("table"), emptyParagraph("separator"), paragraph("body")];

    expect(computeKeepNextChains(blocks).get(1)).toEqual({
      startIndex: 1,
      endIndex: 1,
      memberIndices: [1],
      anchorIndex: 2,
    });
  });

  test("does not treat an authored blank after a table as a structural separator", () => {
    const blocks: FlowBlock[] = [
      table("table"),
      authoredEmptyParagraph("answer space"),
      paragraph("body"),
    ];

    expect(computeKeepNextChains(blocks)).toEqual(new Map());
  });

  test("does not implicitly keep an ordinary empty paragraph with the next paragraph", () => {
    const blocks: FlowBlock[] = [
      paragraph("body before"),
      emptyParagraph("separator"),
      paragraph("body after"),
    ];

    expect(computeKeepNextChains(blocks)).toEqual(new Map());
  });
});

describe("keep-with-next pagination", () => {
  test("keeps a spaced heading on the page when it and the anchor's first line fit", () => {
    const bodySpacing = { before: 0, after: 4 };
    const headingSpacing = { before: 24, after: 8 };
    const anchorSpacing = { before: 8, after: 8 };
    const body = spacedParagraph("body", bodySpacing);
    const heading = spacedParagraph("heading", headingSpacing, true);
    const anchor = spacedParagraph("anchor", anchorSpacing);
    const blocks: FlowBlock[] = [body, heading, anchor];
    const measures: Measure[] = [
      spacedParagraphMeasure(bodySpacing, 800),
      spacedParagraphMeasure(headingSpacing, 24),
      spacedParagraphMeasure(anchorSpacing, 16, 16),
    ];
    // 100 of the 900-unit body remain after the first paragraph. The heading
    // and its two-line anchor need 24 + 24 + 8 + 16 + 16 = 88.
    const result = layoutDocument(blocks, measures, pageOptions);

    const firstPageIds = result.pages[0]?.fragments.map((fragment) => fragment.blockId);
    expect(firstPageIds).toEqual(["body", "heading", "anchor"]);
  });

  test("moves a heading with a three-line widow-controlled paragraph", () => {
    const blocks: FlowBlock[] = [
      paragraph("body"),
      paragraph("heading", true),
      paragraph("anchor"),
    ];
    // 100 units remain after the body: enough for the heading and two anchor
    // lines, but widow control cannot split a three-line paragraph.
    const measures: Measure[] = [
      paragraphMeasure(800),
      paragraphMeasure(40),
      paragraphMeasure(25, 25, 25),
    ];

    const result = layoutDocument(blocks, measures, pageOptions);

    expect(pageBlockIds(result)).toEqual([["body"], ["heading", "anchor"]]);
    expect(result.pages[1]?.fragments.at(-1)).toMatchObject({ fromLine: 0, toLine: 3 });
  });
});

describe("keepLines pagination", () => {
  const keepLinesParagraph = (id: string): ParagraphBlock => ({
    ...paragraph(id),
    attrs: { keepLines: true, widowControl: false },
  });

  test("moves a keepLines paragraph that does not fit whole to the next page", () => {
    const blocks: FlowBlock[] = [paragraph("body"), keepLinesParagraph("kept")];
    const measures: Measure[] = [paragraphMeasure(850), paragraphMeasure(20, 20, 20, 20)];

    const result = layoutDocument(blocks, measures, pageOptions);

    expect(pageBlockIds(result)).toEqual([["body"], ["kept"]]);
    expect(result.pages[1]?.fragments[0]).toMatchObject({ fromLine: 0, toLine: 4 });
  });

  test("splits a keepLines paragraph taller than a full page", () => {
    const blocks: FlowBlock[] = [paragraph("body"), keepLinesParagraph("kept")];
    const measures: Measure[] = [
      paragraphMeasure(850),
      paragraphMeasure(...Array.from({ length: 50 }, () => 20)),
    ];

    const result = layoutDocument(blocks, measures, pageOptions);

    expect(pageBlockIds(result)[0]).toEqual(["body", "kept"]);
    expect(result.pages[0]?.fragments[1]).toMatchObject({ fromLine: 0, toLine: 2 });
  });

  test("splits a paragraph without keepLines at the page end", () => {
    const blocks: FlowBlock[] = [
      paragraph("body"),
      { ...paragraph("plain"), attrs: { widowControl: false } },
    ];
    const measures: Measure[] = [paragraphMeasure(850), paragraphMeasure(20, 20, 20, 20)];

    const result = layoutDocument(blocks, measures, pageOptions);

    expect(pageBlockIds(result)).toEqual([["body", "plain"], ["plain"]]);
  });
});
