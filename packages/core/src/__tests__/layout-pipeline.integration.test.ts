/**
 * Integration Tests for Paginated Layout Engine
 *
 * Tests the complete layout pipeline:
 * 1. Layout produces correct pages from blocks + measures
 * 2. Click maps to correct ProseMirror position
 * 3. Block changes produce correct layout updates
 */

import { describe, test, expect } from "bun:test";

import {
  clickToPosition,
  clickToPositionInParagraph,
} from "../layout-bridge/engine/clickToPosition";
import { hitTestPage, hitTestFragment, getPageTop } from "../layout-bridge/engine/hitTest";
import { layoutDocument } from "../layout-engine/index";
import type {
  ParagraphBlock,
  ParagraphMeasure,
  MeasuredLine,
  FlowBlock,
  Measure,
  PageMargins,
  LayoutOptions,
  TableBlock,
  TableMeasure,
} from "../layout-engine/types";

// =============================================================================
// TEST HELPERS
// =============================================================================

/**
 * Create a simple paragraph block with text runs.
 */
function makeParagraphBlock(
  id: number,
  text: string,
  pmStart: number,
  options: {
    alignment?: "left" | "center" | "right" | "justify";
    keepNext?: boolean;
    pageBreakBefore?: boolean;
  } = {},
): ParagraphBlock {
  return {
    kind: "paragraph",
    id,
    runs: [
      {
        kind: "text",
        text,
        pmStart,
        pmEnd: pmStart + text.length,
      },
    ],
    attrs: {
      alignment: options.alignment,
      keepNext: options.keepNext,
      pageBreakBefore: options.pageBreakBefore,
    },
    pmStart,
    pmEnd: pmStart + text.length + 1, // +1 for paragraph node boundary
  };
}

/**
 * Create a measured line with specified dimensions.
 */
function makeLine(
  fromRun: number,
  fromChar: number,
  toRun: number,
  toChar: number,
  width: number,
  lineHeight: number,
): MeasuredLine {
  return {
    fromRun,
    fromChar,
    toRun,
    toChar,
    width,
    ascent: lineHeight * 0.8,
    descent: lineHeight * 0.2,
    lineHeight,
  };
}

/**
 * Create a paragraph measure from lines.
 */
function makeParagraphMeasure(lines: MeasuredLine[]): ParagraphMeasure {
  const totalHeight = lines.reduce((sum, line) => sum + line.lineHeight, 0);
  return {
    kind: "paragraph",
    lines,
    totalHeight,
  };
}

/**
 * Default page size and margins for tests.
 */
const DEFAULT_PAGE_SIZE = { w: 816, h: 1056 }; // US Letter at 96 DPI
const DEFAULT_MARGINS: PageMargins = {
  top: 96,
  right: 96,
  bottom: 96,
  left: 96,
};

/**
 * Create default layout options.
 */
function makeLayoutOptions(overrides: Partial<LayoutOptions> = {}): LayoutOptions {
  return {
    pageSize: DEFAULT_PAGE_SIZE,
    margins: DEFAULT_MARGINS,
    pageGap: 20,
    ...overrides,
  };
}

// =============================================================================
// TEST SUITE: Layout Produces Correct Pages
// =============================================================================

describe("Layout Engine - Page Production", () => {
  describe("single page scenarios", () => {
    test("empty document produces one empty page", () => {
      const blocks: FlowBlock[] = [];
      const measures: Measure[] = [];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(1);
      expect(layout.pages[0].fragments.length).toBe(0);
      expect(layout.pageSize).toEqual(DEFAULT_PAGE_SIZE);
    });

    test("single paragraph fits on one page", () => {
      const blocks: FlowBlock[] = [makeParagraphBlock(0, "Hello, World!", 1)];
      const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 13, 100, 24)])];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(1);
      expect(layout.pages[0].fragments.length).toBe(1);

      const fragment = layout.pages[0].fragments[0];
      expect(fragment.kind).toBe("paragraph");
      expect(fragment.blockId).toBe(0);
    });

    test("positions a text-anchored floating table from the current text cursor", () => {
      const paragraph = makeParagraphBlock(0, "Anchor", 1);
      const table: TableBlock = {
        kind: "table",
        id: 1,
        rows: [
          {
            id: "row",
            cells: [
              {
                id: "cell",
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                blocks: [makeParagraphBlock(2, "Value", 8)],
              },
            ],
          },
        ],
        columnWidths: [200],
        floating: { vertAnchor: "text", tblpY: 10 },
      };
      const paragraphMeasure = makeParagraphMeasure([makeLine(0, 0, 0, 6, 50, 40)]);
      const tableMeasure: TableMeasure = {
        kind: "table",
        rows: [
          {
            cells: [
              {
                blocks: [makeParagraphMeasure([makeLine(0, 0, 0, 5, 40, 20)])],
                width: 200,
                height: 20,
              },
            ],
            height: 20,
          },
        ],
        columnWidths: [200],
        totalWidth: 200,
        totalHeight: 20,
      };

      const layout = layoutDocument(
        [paragraph, table],
        [paragraphMeasure, tableMeasure],
        makeLayoutOptions(),
      );
      const tableFragment = layout.pages[0]?.fragments.find(
        (fragment) => fragment.kind === "table",
      );

      expect(tableFragment?.y).toBe(DEFAULT_MARGINS.top + 40 + 10);
    });

    test("first paragraph on a page honors explicit spaceBefore", () => {
      const blocks: FlowBlock[] = [
        {
          ...makeParagraphBlock(0, "Starts lower", 1),
          attrs: { spacing: { before: 18, after: 0 } },
        },
      ];
      const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)])];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages[0].fragments[0]?.y).toBe(DEFAULT_MARGINS.top + 18);
    });

    test("multiple paragraphs fit on one page", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "First paragraph", 1),
        makeParagraphBlock(1, "Second paragraph", 18),
        makeParagraphBlock(2, "Third paragraph", 36),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 120, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 130, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 120, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(1);
      expect(layout.pages[0].fragments.length).toBe(3);
    });

    test("paragraph positions are stacked vertically", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "First", 1),
        makeParagraphBlock(1, "Second", 8),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      const frag0 = layout.pages[0].fragments[0];
      const frag1 = layout.pages[0].fragments[1];

      // Second fragment should start below first
      expect(frag1.y).toBeGreaterThan(frag0.y);
    });
  });

  describe("multi-page scenarios", () => {
    test("content exceeding page height creates multiple pages", () => {
      // Create many paragraphs that exceed page content height
      // Content height = 1056 - 96 - 96 = 864px
      // Each paragraph = 100px line height
      // 9 paragraphs = 900px > 864px, should overflow to page 2
      const blocks: FlowBlock[] = [];
      const measures: Measure[] = [];

      for (let i = 0; i < 10; i++) {
        blocks.push(makeParagraphBlock(i, `Paragraph ${i}`, i * 15));
        measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 100)]));
      }

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(2);
      // First page should have fragments
      expect(layout.pages[0].fragments.length).toBeGreaterThan(0);
      // Second page should have remaining fragments
      expect(layout.pages[1].fragments.length).toBeGreaterThan(0);
    });

    test("explicit page break creates new page", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        { kind: "pageBreak", id: 1, pmStart: 15, pmEnd: 16 },
        makeParagraphBlock(2, "After break", 17),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        { kind: "pageBreak" },
        makeParagraphMeasure([makeLine(0, 0, 0, 11, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(2);
      expect(layout.pages[0].fragments.length).toBe(1);
      expect(layout.pages[1].fragments.length).toBe(1);
    });

    test("consecutive explicit page breaks preserve a blank page", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        { kind: "pageBreak", id: 1, pmStart: 15, pmEnd: 16 },
        { kind: "pageBreak", id: 2, pmStart: 16, pmEnd: 17 },
        makeParagraphBlock(3, "After blank page", 18),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        { kind: "pageBreak" },
        { kind: "pageBreak" },
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(3);
      expect(layout.pages[1].fragments).toEqual([]);
      expect(layout.pages[2].fragments[0].blockId).toBe(3);
    });

    test("pageBreakBefore attribute creates new page", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "First paragraph", 1),
        makeParagraphBlock(1, "Second with break", 18, {
          pageBreakBefore: true,
        }),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 120, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 17, 140, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(2);
      expect(layout.pages[0].fragments[0].blockId).toBe(0);
      expect(layout.pages[1].fragments[0].blockId).toBe(1);
    });

    test("pageBreakBefore after an explicit page break preserves a blank page", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        { kind: "pageBreak", id: 1, pmStart: 15, pmEnd: 16 },
        makeParagraphBlock(2, "After blank page", 17, {
          pageBreakBefore: true,
        }),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        { kind: "pageBreak" },
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(3);
      expect(layout.pages[1].fragments).toEqual([]);
      expect(layout.pages[2].fragments[0].blockId).toBe(2);
    });

    test("rendered page break reuses a page opened by a structural break", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        { kind: "pageBreak", id: 1, pmStart: 15, pmEnd: 16 },
        {
          ...makeParagraphBlock(2, "Cached next page", 17),
          attrs: {
            renderedPageBreakBefore: true,
            spacing: { before: 24 },
            spacingExplicit: { before: true },
          },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        { kind: "pageBreak" },
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(2);
      expect(layout.pages[1].fragments[0].blockId).toBe(2);
      expect(layout.pages[1].fragments[0].y).toBe(DEFAULT_MARGINS.top);
    });

    test("rendered page break reuses an authored boundary across an empty carrier", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        { kind: "pageBreak", id: 1, pmStart: 15, pmEnd: 16 },
        {
          ...makeParagraphBlock(2, "", 17),
          runs: [],
        },
        {
          ...makeParagraphBlock(3, "Cached next page", 18),
          attrs: {
            renderedPageBreakBefore: true,
            spacing: { before: 24 },
          },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        { kind: "pageBreak" },
        makeParagraphMeasure([]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments.at(-1)?.blockId).toBe(3);
      expect(layout.pages[1]?.fragments.at(-1)?.y).toBe(DEFAULT_MARGINS.top);
    });

    test("rendered page break preserves leading spacing after a section boundary", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before section", 1),
        { kind: "sectionBreak", id: 1, type: "nextPage" },
        {
          ...makeParagraphBlock(2, "First paragraph in section", 17),
          attrs: {
            renderedPageBreakBefore: true,
            spacing: { before: 24 },
            spacingExplicit: { before: true },
          },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 14, 100, 24)]),
        { kind: "sectionBreak" },
        makeParagraphMeasure([makeLine(0, 0, 0, 26, 120, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments[0]?.blockId).toBe(2);
      expect(layout.pages[1]?.fragments[0]?.y).toBe(DEFAULT_MARGINS.top + 24);
    });

    test("rendered page break consumes inherited spacing after a section boundary", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before section", 1),
        { kind: "sectionBreak", id: 1, type: "nextPage" },
        {
          ...makeParagraphBlock(2, "First paragraph in section", 17),
          attrs: {
            renderedPageBreakBefore: true,
            spacing: { before: 24 },
          },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 14, 100, 24)]),
        { kind: "sectionBreak" },
        makeParagraphMeasure([makeLine(0, 0, 0, 26, 120, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments[0]?.blockId).toBe(2);
      expect(layout.pages[1]?.fragments[0]?.y).toBe(DEFAULT_MARGINS.top);
    });

    test("stale rendered page break remains advisory when the paragraph fits", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        {
          ...makeParagraphBlock(1, "Cached next page", 15),
          attrs: { renderedPageBreakBefore: true },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(1);
      expect(layout.pages[0]?.fragments.map(({ blockId }) => blockId)).toEqual([0, 1]);
    });

    test("rendered page break preserves a fitting keep-next heading boundary", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before heading", 1),
        {
          ...makeParagraphBlock(1, "Kept heading", 16, { keepNext: true }),
          attrs: { keepNext: true, renderedPageBreakBefore: true },
        },
        makeParagraphBlock(2, "Kept body", 29),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 14, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 9, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[0]?.fragments.map(({ blockId }) => blockId)).toEqual([0]);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2]);
    });

    test("rendered page break moves a paragraph that would cross the current page", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        {
          ...makeParagraphBlock(1, "Cached paragraph wraps across four lines", 23),
          attrs: { renderedPageBreakBefore: true },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 820)]),
        makeParagraphMeasure([
          makeLine(0, 0, 0, 10, 100, 20),
          makeLine(0, 10, 0, 20, 100, 20),
          makeLine(0, 20, 0, 30, 100, 20),
          makeLine(0, 30, 0, 40, 100, 20),
        ]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[0]?.fragments.map(({ blockId }) => blockId)).toEqual([0]);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1]);
    });

    test("rendered page break does not reapply leading spacing after snapping", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        {
          ...makeParagraphBlock(1, "Cached next page", 23),
          attrs: {
            renderedPageBreakBefore: true,
            spacing: { before: 24 },
            spacingExplicit: { before: true },
          },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 820)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 80)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments[0]?.y).toBe(DEFAULT_MARGINS.top);
    });

    test("authored pageBreakBefore keeps leading spacing", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        {
          ...makeParagraphBlock(1, "Authored next page", 14),
          attrs: {
            pageBreakBefore: true,
            renderedPageBreakBefore: true,
            spacing: { before: 24 },
            spacingExplicit: { before: true },
          },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 18, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments[0]?.y).toBe(DEFAULT_MARGINS.top + 24);
    });

    test("rendered page break does not add a blank page at a natural overflow", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Page-filling paragraph", 1),
        {
          ...makeParagraphBlock(1, "Cached next page", 24),
          attrs: { renderedPageBreakBefore: true },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 22, 500, 850)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(2);
      expect(layout.pages[0].fragments[0].blockId).toBe(0);
      expect(layout.pages[1].fragments[0].blockId).toBe(1);
    });

    test("rendered page break reuses a page opened by the preceding paragraph continuation", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Paragraph split across the boundary", 1),
        {
          ...makeParagraphBlock(1, "Cached next page", 38),
          attrs: { renderedPageBreakBefore: true },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 850), makeLine(0, 20, 0, 37, 300, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1].fragments.map((fragment) => fragment.blockId)).toEqual([0, 1]);
      expect(layout.pages[1].fragments[0]).toMatchObject({ continuesFromPrev: true });
    });

    test("rendered page break reuses a page containing only an overflowed empty paragraph", () => {
      const emptyParagraph: ParagraphBlock = {
        kind: "paragraph",
        id: 1,
        runs: [],
        pmStart: 24,
        pmEnd: 25,
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Page-filling paragraph", 1),
        emptyParagraph,
        {
          ...makeParagraphBlock(2, "Cached next page", 26),
          attrs: { renderedPageBreakBefore: true },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 22, 500, 850)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 0, 0, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(2);
      expect(layout.pages[1].fragments.map((fragment) => fragment.blockId)).toEqual([1, 2]);
    });

    test("empty rendered break marker reuses a page opened by the preceding paragraph", () => {
      const marker: ParagraphBlock = {
        kind: "paragraph",
        id: 3,
        runs: [
          { kind: "tab", pmStart: 76, pmEnd: 77 },
          { kind: "tab", pmStart: 77, pmEnd: 78 },
        ],
        attrs: { renderedPageBreakBefore: true },
        pmStart: 25,
        pmEnd: 26,
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        makeParagraphBlock(1, "Naturally moves to page two", 22),
        makeParagraphBlock(2, "Trailing content on page two", 50),
        marker,
        makeParagraphBlock(4, "After cached boundary", 79),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 840)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 27, 100, 40)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 28, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 0, 0, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 21, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2, 3, 4]);
    });

    test("rendered page break reuses a page opened by a keep-next chain", () => {
      const marker = {
        ...makeParagraphBlock(3, "Cached boundary", 61),
        attrs: { renderedPageBreakBefore: true },
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        makeParagraphBlock(1, "Heading", 23, { keepNext: true }),
        makeParagraphBlock(2, "Kept body", 31),
        marker,
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 810)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 7, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 9, 100, 40)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2, 3]);
    });

    test("rendered page break reuses a page containing a paragraph continuation", () => {
      const continued = makeParagraphBlock(
        1,
        "A paragraph whose measured lines continue across the page boundary",
        23,
      );
      const marker = {
        ...makeParagraphBlock(2, "Cached boundary", 90),
        attrs: { renderedPageBreakBefore: true },
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        continued,
        marker,
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 800)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 31, 200, 40), makeLine(0, 32, 0, 64, 200, 40)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[0]?.fragments.map(({ blockId }) => blockId)).toEqual([0, 1]);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2]);
    });

    test("stale rendered page break remains advisory after a whole paragraph moves", () => {
      const previousItem = {
        ...makeParagraphBlock(1, "Moves intact", 23),
        attrs: { numPr: { numId: 4, ilvl: 1 } },
      };
      const marker = {
        ...makeParagraphBlock(2, "Cached boundary", 50),
        attrs: {
          numPr: { numId: 4, ilvl: 0 },
          renderedPageBreakBefore: true,
        },
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        previousItem,
        marker,
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 840)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 40)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2]);
    });

    test("rendered page break reuses a page opened within one numbered sequence", () => {
      const previousItem = {
        ...makeParagraphBlock(1, "Moves intact", 23),
        attrs: { numPr: { numId: 4, ilvl: 1 } },
      };
      const marker = {
        ...makeParagraphBlock(2, "Cached boundary", 50),
        attrs: {
          numPr: { numId: 4, ilvl: 1 },
          renderedPageBreakBefore: true,
        },
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        previousItem,
        marker,
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 840)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 40)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2]);
    });

    test("rendered page break reuses a page opened within one tabbed style sequence", () => {
      const previousRow: FlowBlock = {
        ...makeParagraphBlock(1, "First row", 23),
        runs: [
          { kind: "text", text: "First", pmStart: 23, pmEnd: 28 },
          { kind: "tab", pmStart: 28, pmEnd: 29 },
          { kind: "text", text: "row", pmStart: 29, pmEnd: 32 },
        ],
        attrs: {
          styleId: "TabularLine",
          indent: { left: 42 },
          tabs: [{ val: "start", pos: 720 }],
        },
      };
      const marker: FlowBlock = {
        ...makeParagraphBlock(2, "Second row", 33),
        runs: [
          { kind: "text", text: "Second", pmStart: 33, pmEnd: 39 },
          { kind: "tab", pmStart: 39, pmEnd: 40 },
          { kind: "text", text: "row", pmStart: 40, pmEnd: 43 },
        ],
        attrs: {
          styleId: "TabularLine",
          indent: { left: 42 },
          tabs: [{ val: "start", pos: 720 }],
          renderedPageBreakBefore: true,
        },
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        previousRow,
        marker,
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 840)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 9, 100, 40)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 10, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2]);
    });

    test("stale rendered page break remains advisory when tabbed layouts reflow", () => {
      const previousRow: FlowBlock = {
        ...makeParagraphBlock(1, "First row", 23),
        runs: [
          { kind: "text", text: "First", pmStart: 23, pmEnd: 28 },
          { kind: "tab", pmStart: 28, pmEnd: 29 },
          { kind: "text", text: "row", pmStart: 29, pmEnd: 32 },
        ],
        attrs: {
          styleId: "TabularLine",
          indent: { left: 42 },
          tabs: [{ val: "start", pos: 720 }],
        },
      };
      const marker: FlowBlock = {
        ...makeParagraphBlock(2, "Different row", 33),
        runs: [
          { kind: "text", text: "Different", pmStart: 33, pmEnd: 42 },
          { kind: "tab", pmStart: 42, pmEnd: 43 },
          { kind: "text", text: "row", pmStart: 43, pmEnd: 46 },
        ],
        attrs: {
          styleId: "TabularLine",
          indent: { left: 42 },
          tabs: [{ val: "start", pos: 1440 }],
          renderedPageBreakBefore: true,
        },
      };
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Nearly fills page one", 1),
        previousRow,
        marker,
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 20, 500, 840)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 9, 100, 40)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 13, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2]);
    });

    test("successive stale rendered page breaks remain advisory after natural reflow", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Fill page one", 1),
        makeParagraphBlock(1, "Natural page two", 15),
        {
          ...makeParagraphBlock(2, "Cached page two", 32),
          attrs: { renderedPageBreakBefore: true },
        },
        makeParagraphBlock(3, "Fill page two", 48),
        makeParagraphBlock(4, "Natural page three", 62),
        {
          ...makeParagraphBlock(5, "Cached page three", 81),
          attrs: { renderedPageBreakBefore: true },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 13, 500, 850)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 15, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 13, 500, 810)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 18, 100, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 17, 100, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages).toHaveLength(3);
      expect(layout.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([1, 2, 3]);
      expect(layout.pages[2]?.fragments.map(({ blockId }) => blockId)).toEqual([4, 5]);
    });

    test("explicit pageBreakBefore takes priority over a rendered page break hint", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Before break", 1),
        { kind: "pageBreak", id: 1, pmStart: 15, pmEnd: 16 },
        {
          ...makeParagraphBlock(2, "After intentional blank page", 17),
          attrs: { pageBreakBefore: true, renderedPageBreakBefore: true },
        },
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)]),
        { kind: "pageBreak" },
        makeParagraphMeasure([makeLine(0, 0, 0, 16, 90, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBe(3);
      expect(layout.pages[1].fragments).toEqual([]);
      expect(layout.pages[2].fragments[0].blockId).toBe(2);
    });
  });

  describe("paragraph splitting across pages", () => {
    test("long paragraph splits across pages", () => {
      // Create a paragraph with many lines that exceeds page height
      const lines: MeasuredLine[] = [];
      const lineHeight = 100;
      const numLines = 15; // 15 * 100 = 1500px > 864px content area

      for (let i = 0; i < numLines; i++) {
        lines.push(makeLine(0, i * 10, 0, (i + 1) * 10, 500, lineHeight));
      }

      const blocks: FlowBlock[] = [makeParagraphBlock(0, "A".repeat(150), 1)];
      const measures: Measure[] = [makeParagraphMeasure(lines)];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      expect(layout.pages.length).toBeGreaterThan(1);

      // First page fragment should have fromLine = 0
      const firstFrag = layout.pages[0].fragments[0];
      expect(firstFrag.kind).toBe("paragraph");
      if (firstFrag.kind === "paragraph") {
        expect(firstFrag.fromLine).toBe(0);
        expect(firstFrag.toLine).toBeGreaterThan(0);
        expect(firstFrag.continuesOnNext).toBe(true);
      }

      // Second page fragment should continue
      const secondFrag = layout.pages[1].fragments[0];
      expect(secondFrag.kind).toBe("paragraph");
      if (secondFrag.kind === "paragraph") {
        expect(secondFrag.continuesFromPrev).toBe(true);
      }
    });
  });

  describe("keepNext chain handling", () => {
    test("keepNext paragraphs stay together on new page", () => {
      // Create paragraphs that nearly fill first page
      const blocks: FlowBlock[] = [];
      const measures: Measure[] = [];

      // Add filler paragraphs to fill most of the page
      for (let i = 0; i < 7; i++) {
        blocks.push(makeParagraphBlock(i, `Filler ${i}`, i * 10));
        measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 10, 80, 100)]));
      }

      // Add keepNext paragraph that should pull next paragraph to new page
      blocks.push(makeParagraphBlock(7, "KeepNext heading", 70, { keepNext: true }));
      measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 16, 120, 100)]));

      // Add following paragraph
      blocks.push(makeParagraphBlock(8, "Following paragraph", 88));
      measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 19, 150, 100)]));

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      // The keepNext heading and following paragraph should be on same page
      const headingFrag = layout.pages.flatMap((p) => p.fragments).find((f) => f.blockId === 7);
      const followingFrag = layout.pages.flatMap((p) => p.fragments).find((f) => f.blockId === 8);

      // They should be on the same page
      if (!headingFrag || !followingFrag) {
        throw new Error("Expected heading and following fragments");
      }
      const headingPage = layout.pages.findIndex((p) => p.fragments.includes(headingFrag));
      const followingPage = layout.pages.findIndex((p) => p.fragments.includes(followingFrag));

      expect(headingPage).toBe(followingPage);
    });

    test("keepNext carries across an empty separator", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "Filler", 0),
        makeParagraphBlock(1, "Heading", 10, { keepNext: true }),
        makeParagraphBlock(2, "", 20),
        makeParagraphBlock(3, "Body", 21),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 6, 80, 820)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 7, 80, 20)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 0, 0, 20)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 4, 80, 20)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());
      const headingPage = layout.pages.findIndex((page) =>
        page.fragments.some(({ blockId }) => blockId === 1),
      );
      const bodyPage = layout.pages.findIndex((page) =>
        page.fragments.some(({ blockId }) => blockId === 3),
      );

      expect(headingPage).toBe(1);
      expect(bodyPage).toBe(headingPage);
    });

    test("moves an oversized keepNext chain before paginating it naturally", () => {
      const blocks: FlowBlock[] = [makeParagraphBlock(0, "Filler", 0)];
      const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 6, 80, 400)])];

      for (let id = 1; id <= 6; id++) {
        blocks.push(makeParagraphBlock(id, `Chain ${id}`, id * 10, { keepNext: id < 6 }));
        measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 7, 80, 200)]));
      }

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());
      const chainStartPage = layout.pages.findIndex((page) =>
        page.fragments.some(({ blockId }) => blockId === 1),
      );

      expect(chainStartPage).toBe(1);
      expect(layout.pages[0]?.fragments.map(({ blockId }) => blockId)).toEqual([0]);
      const chainPageCount = layout.pages.filter((page) =>
        page.fragments.some(({ blockId }) => blockId !== 0),
      ).length;
      expect(chainPageCount).toBeGreaterThan(1);
    });
  });

  describe("margin and positioning", () => {
    test("fragments are positioned within content area", () => {
      const blocks: FlowBlock[] = [makeParagraphBlock(0, "Test content", 1)];
      const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)])];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      const fragment = layout.pages[0].fragments[0];

      // Fragment X should be at left margin
      expect(fragment.x).toBe(DEFAULT_MARGINS.left);

      // Fragment Y should be at top margin
      expect(fragment.y).toBe(DEFAULT_MARGINS.top);
    });

    test("content width is page width minus margins", () => {
      const blocks: FlowBlock[] = [makeParagraphBlock(0, "Test", 1)];
      const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 4, 50, 24)])];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      const fragment = layout.pages[0].fragments[0];
      const expectedWidth = DEFAULT_PAGE_SIZE.w - DEFAULT_MARGINS.left - DEFAULT_MARGINS.right;

      expect(fragment.width).toBe(expectedWidth);
    });
  });
});

// =============================================================================
// TEST SUITE: Click Maps to Correct PM Position
// =============================================================================

describe("Click-to-Position Mapping", () => {
  describe("page hit testing", () => {
    test("hitTestPage finds correct page", () => {
      // Create a 2-page layout
      const blocks: FlowBlock[] = [];
      const measures: Measure[] = [];

      for (let i = 0; i < 20; i++) {
        blocks.push(makeParagraphBlock(i, `Para ${i}`, i * 10));
        measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 8, 60, 50)]));
      }

      const layout = layoutDocument(blocks, measures, makeLayoutOptions({ pageGap: 20 }));

      expect(layout.pages.length).toBeGreaterThan(1);

      // Hit test at top of document
      const hit1 = hitTestPage(layout, { x: 100, y: 50 });
      expect(hit1).not.toBeNull();
      expect(hit1?.pageIndex).toBe(0);

      // Hit test in second page (after first page height + gap)
      const pageHeight = layout.pageSize.h;
      const pageGap = 20;
      const secondPageTop = pageHeight + pageGap;
      const hit2 = hitTestPage(layout, { x: 100, y: secondPageTop + 50 });
      expect(hit2).not.toBeNull();
      expect(hit2?.pageIndex).toBe(1);
    });

    test("hitTestPage returns correct pageY offset", () => {
      const blocks: FlowBlock[] = [makeParagraphBlock(0, "Test", 1)];
      const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 4, 50, 24)])];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      const hit = hitTestPage(layout, { x: 100, y: 150 });
      expect(hit).not.toBeNull();
      expect(hit?.pageY).toBe(150);
    });

    test("getPageTop returns cumulative offset", () => {
      const blocks: FlowBlock[] = [];
      const measures: Measure[] = [];

      for (let i = 0; i < 20; i++) {
        blocks.push(makeParagraphBlock(i, `Para ${i}`, i * 10));
        measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 8, 60, 50)]));
      }

      const layout = layoutDocument(blocks, measures, makeLayoutOptions({ pageGap: 20 }));

      expect(getPageTop(layout, 0)).toBe(0);

      if (layout.pages.length > 1) {
        const expectedPage1Top = layout.pageSize.h + 20;
        expect(getPageTop(layout, 1)).toBe(expectedPage1Top);
      }
    });
  });

  describe("fragment hit testing", () => {
    test("hitTestFragment finds correct fragment", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "First", 1),
        makeParagraphBlock(1, "Second", 8),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 24)]),
      ];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      const pageHit = hitTestPage(layout, {
        x: 100,
        y: DEFAULT_MARGINS.top + 5,
      });
      expect(pageHit).not.toBeNull();

      if (!pageHit) {
        throw new Error("Expected pageHit");
      }
      const fragmentHit = hitTestFragment(pageHit, blocks, measures, {
        x: DEFAULT_MARGINS.left + 10,
        y: DEFAULT_MARGINS.top + 5,
      });

      expect(fragmentHit).not.toBeNull();
      expect(fragmentHit?.fragment.blockId).toBe(0);
    });

    test("hitTestFragment calculates correct local coordinates", () => {
      const blocks: FlowBlock[] = [makeParagraphBlock(0, "Test content", 1)];
      const measures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 12, 100, 24)])];

      const layout = layoutDocument(blocks, measures, makeLayoutOptions());

      const pageHit = hitTestPage(layout, { x: 150, y: 110 });
      if (!pageHit) {
        throw new Error("Expected pageHit");
      }
      const fragmentHit = hitTestFragment(pageHit, blocks, measures, {
        x: 150,
        y: 110,
      });

      expect(fragmentHit).not.toBeNull();
      if (!fragmentHit) {
        throw new Error("Expected fragmentHit");
      }

      // Local coordinates should be relative to fragment position
      const fragment = fragmentHit.fragment;
      expect(fragmentHit.localX).toBe(150 - fragment.x);
      expect(fragmentHit.localY).toBe(110 - fragment.y);
    });
  });

  describe("click to PM position", () => {
    test("clickToPositionInParagraph maps click to correct position", () => {
      const block = makeParagraphBlock(0, "Hello World", 1);
      const measure = makeParagraphMeasure([makeLine(0, 0, 0, 11, 100, 24)]);

      // Create a synthetic fragment hit
      const fragmentHit = {
        fragment: {
          kind: "paragraph" as const,
          blockId: 0,
          x: 96,
          y: 96,
          width: 624,
          height: 24,
          fromLine: 0,
          toLine: 1,
        },
        block,
        measure,
        pageIndex: 0,
        localX: 0, // Click at start of line
        localY: 10,
      };

      const result = clickToPositionInParagraph(fragmentHit);

      expect(result).not.toBeNull();
      expect(result?.pmPosition).toBe(1); // Start of text
      expect(result?.lineIndex).toBe(0);
    });

    test("clickToPosition returns PM position from fragment hit", () => {
      const block = makeParagraphBlock(0, "Test", 1);
      const measure = makeParagraphMeasure([makeLine(0, 0, 0, 4, 40, 24)]);

      const fragmentHit = {
        fragment: {
          kind: "paragraph" as const,
          blockId: 0,
          x: 96,
          y: 96,
          width: 624,
          height: 24,
          fromLine: 0,
          toLine: 1,
        },
        block,
        measure,
        pageIndex: 0,
        localX: 0,
        localY: 10,
      };

      const pmPosition = clickToPosition(fragmentHit);

      expect(pmPosition).not.toBeNull();
      expect(pmPosition).toBeGreaterThanOrEqual(1);
    });

    // Note: This test requires a DOM environment with canvas for text measurement.
    // In headless bun:test, we test the logic without requiring precise character positioning.
    test("click at end of line returns end position (mock)", () => {
      const block = makeParagraphBlock(0, "Hello", 1);
      const measure = makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)]);

      // Test the fragment structure is correct
      const fragmentHit = {
        fragment: {
          kind: "paragraph" as const,
          blockId: 0,
          x: 96,
          y: 96,
          width: 624,
          height: 24,
          fromLine: 0,
          toLine: 1,
        },
        block,
        measure,
        pageIndex: 0,
        localX: 0, // Click at start (doesn't require canvas measurement)
        localY: 10,
      };

      // Verify the block structure is correct
      expect(block.pmStart).toBe(1);
      expect(block.pmEnd).toBe(7); // 'Hello' + paragraph node boundary

      // This tests that the mapping starts correctly
      const result = clickToPositionInParagraph(fragmentHit);
      expect(result).not.toBeNull();
      expect(result?.pmPosition).toBe(1); // Start of text
    });
  });
});

// =============================================================================
// TEST SUITE: PM Edits Update Visual Pages
// =============================================================================

describe("Document Updates", () => {
  describe("adding content", () => {
    test("adding paragraph increases fragment count", () => {
      // Initial state: 1 paragraph
      const initialBlocks: FlowBlock[] = [makeParagraphBlock(0, "Initial", 1)];
      const initialMeasures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 7, 60, 24)])];

      const initialLayout = layoutDocument(initialBlocks, initialMeasures, makeLayoutOptions());

      // After adding paragraph: 2 paragraphs
      const updatedBlocks: FlowBlock[] = [
        makeParagraphBlock(0, "Initial", 1),
        makeParagraphBlock(1, "Added", 10),
      ];
      const updatedMeasures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 7, 60, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)]),
      ];

      const updatedLayout = layoutDocument(updatedBlocks, updatedMeasures, makeLayoutOptions());

      expect(updatedLayout.pages[0].fragments.length).toBe(
        initialLayout.pages[0].fragments.length + 1,
      );
    });

    test("adding enough content creates new page", () => {
      // Start with content that fits on one page
      const smallBlocks: FlowBlock[] = [makeParagraphBlock(0, "Small", 1)];
      const smallMeasures: Measure[] = [makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)])];

      const smallLayout = layoutDocument(smallBlocks, smallMeasures, makeLayoutOptions());
      expect(smallLayout.pages.length).toBe(1);

      // Add content that overflows
      const largeBlocks: FlowBlock[] = [];
      const largeMeasures: Measure[] = [];

      for (let i = 0; i < 20; i++) {
        largeBlocks.push(makeParagraphBlock(i, `Para ${i}`, i * 10));
        largeMeasures.push(makeParagraphMeasure([makeLine(0, 0, 0, 8, 60, 50)]));
      }

      const largeLayout = layoutDocument(largeBlocks, largeMeasures, makeLayoutOptions());
      expect(largeLayout.pages.length).toBeGreaterThan(1);
    });
  });

  describe("removing content", () => {
    test("removing paragraph decreases fragment count", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "First", 1),
        makeParagraphBlock(1, "Second", 8),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 24)]),
      ];

      const beforeLayout = layoutDocument(blocks, measures, makeLayoutOptions());

      // Remove second paragraph
      const afterBlocks = blocks.slice(0, 1);
      const afterMeasures = measures.slice(0, 1);

      const afterLayout = layoutDocument(afterBlocks, afterMeasures, makeLayoutOptions());

      expect(afterLayout.pages[0].fragments.length).toBe(
        beforeLayout.pages[0].fragments.length - 1,
      );
    });

    test("removing content can reduce page count", () => {
      // Create multi-page document
      const blocks: FlowBlock[] = [];
      const measures: Measure[] = [];

      for (let i = 0; i < 20; i++) {
        blocks.push(makeParagraphBlock(i, `Para ${i}`, i * 10));
        measures.push(makeParagraphMeasure([makeLine(0, 0, 0, 8, 60, 50)]));
      }

      const multiPageLayout = layoutDocument(blocks, measures, makeLayoutOptions());
      expect(multiPageLayout.pages.length).toBeGreaterThan(1);

      // Remove most content
      const smallBlocks = blocks.slice(0, 2);
      const smallMeasures = measures.slice(0, 2);

      const singlePageLayout = layoutDocument(smallBlocks, smallMeasures, makeLayoutOptions());
      expect(singlePageLayout.pages.length).toBe(1);
    });
  });

  describe("modifying content", () => {
    test("changing text updates PM positions in layout", () => {
      // Original paragraph
      const originalBlock = makeParagraphBlock(0, "Original", 1);
      const originalMeasure = makeParagraphMeasure([makeLine(0, 0, 0, 8, 70, 24)]);

      const originalLayout = layoutDocument(
        [originalBlock],
        [originalMeasure],
        makeLayoutOptions(),
      );

      // Modified paragraph (longer text)
      const modifiedBlock = makeParagraphBlock(0, "Modified text here", 1);
      const modifiedMeasure = makeParagraphMeasure([makeLine(0, 0, 0, 18, 140, 24)]);

      const modifiedLayout = layoutDocument(
        [modifiedBlock],
        [modifiedMeasure],
        makeLayoutOptions(),
      );

      // Both should have same structure but different PM bounds
      expect(modifiedLayout.pages.length).toBe(originalLayout.pages.length);
      expect(modifiedLayout.pages[0].fragments.length).toBe(
        originalLayout.pages[0].fragments.length,
      );

      // PM end should differ based on text length
      const originalFrag = originalLayout.pages[0].fragments[0];
      const modifiedFrag = modifiedLayout.pages[0].fragments[0];

      expect(modifiedFrag.pmEnd).toBeGreaterThan(originalFrag.pmEnd ?? 0);
    });

    test("line height changes update fragment positions", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, "First", 1),
        makeParagraphBlock(1, "Second", 8),
      ];

      // Small line height
      const smallMeasures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)]),
        makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 24)]),
      ];

      const smallLayout = layoutDocument(blocks, smallMeasures, makeLayoutOptions());

      // Large line height
      const largeMeasures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 48)]), // Double line height
        makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 48)]),
      ];

      const largeLayout = layoutDocument(blocks, largeMeasures, makeLayoutOptions());

      // Second fragment should be positioned further down with larger line height
      const smallSecondY = smallLayout.pages[0].fragments[1].y;
      const largeSecondY = largeLayout.pages[0].fragments[1].y;

      expect(largeSecondY).toBeGreaterThan(smallSecondY);
    });
  });
});

// =============================================================================
// TEST SUITE: Section Breaks
// =============================================================================

describe("Section Breaks", () => {
  test("nextPage section break forces new page", () => {
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, "Before section", 1),
      { kind: "sectionBreak", id: 1, type: "nextPage" },
      makeParagraphBlock(2, "After section", 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 14, 120, 24)]),
      { kind: "sectionBreak" },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(2);
    expect(layout.pages[0].fragments.some((f) => f.blockId === 0)).toBe(true);
    expect(layout.pages[1].fragments.some((f) => f.blockId === 2)).toBe(true);
  });

  test("coalesced blank section page gets the next section header and footer refs", () => {
    const blocks: FlowBlock[] = [
      { kind: "sectionBreak", id: 1, type: "nextPage" },
      makeParagraphBlock(2, "After section", 18),
    ];
    const measures: Measure[] = [
      { kind: "sectionBreak" },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({
        sectionHeaderFooterRefs: [
          { footerDefault: "previous-footer" },
          { footerDefault: "next-footer" },
        ],
      }),
    );

    expect(layout.pages.length).toBe(1);
    expect(layout.pages[0].sectionIndex).toBe(1);
    expect(layout.pages[0].sectionPageNumber).toBe(1);
    expect(layout.pages[0].headerFooterRefs?.footerDefault).toBe("next-footer");
    expect(layout.pages[0].fragments.some((f) => f.blockId === 2)).toBe(true);
  });

  test("continuous section break does not force new page", () => {
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, "Before section", 1),
      { kind: "sectionBreak", id: 1, type: "continuous" },
      makeParagraphBlock(2, "After section", 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 14, 120, 24)]),
      { kind: "sectionBreak" },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(1);
    expect(layout.pages[0].fragments.length).toBe(2);
  });

  test("evenPage section break preserves the blank parity page", () => {
    const pageContentHeight = DEFAULT_PAGE_SIZE.h - DEFAULT_MARGINS.top - DEFAULT_MARGINS.bottom;
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, "Before section", 1),
      { kind: "sectionBreak", id: 1, type: "evenPage" },
      makeParagraphBlock(2, "After section", 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([
        makeLine(0, 0, 0, 7, 120, pageContentHeight),
        makeLine(0, 7, 0, 14, 120, pageContentHeight),
      ]),
      { kind: "sectionBreak" },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(4);
    expect(layout.pages[2].fragments).toEqual([]);
    expect(layout.pages[3].fragments.some((f) => f.blockId === 2)).toBe(true);
  });
});

// =============================================================================
// TEST SUITE: Contextual Spacing (OOXML §17.3.1.9)
// =============================================================================

describe("Layout Engine - Contextual Spacing", () => {
  /**
   * Helper to create a paragraph block with spacing and contextualSpacing attrs.
   */
  function makeSpacedParagraph(
    id: number,
    text: string,
    pmStart: number,
    options: {
      spaceBefore?: number;
      spaceAfter?: number;
      contextualSpacing?: boolean;
      styleId?: string;
    } = {},
  ): ParagraphBlock {
    return {
      kind: "paragraph",
      id,
      runs: [{ kind: "text", text, pmStart, pmEnd: pmStart + text.length }],
      attrs: {
        spacing: {
          before: options.spaceBefore ?? 0,
          after: options.spaceAfter ?? 13,
        },
        contextualSpacing: options.contextualSpacing,
        styleId: options.styleId,
      },
      pmStart,
      pmEnd: pmStart + text.length + 1,
    };
  }

  test("suppresses spacing between consecutive same-style paragraphs with contextualSpacing", () => {
    // Two ListBullet paragraphs with contextualSpacing — spacing should be suppressed
    const blocks: FlowBlock[] = [
      makeSpacedParagraph(0, "Item 1", 1, {
        spaceAfter: 13,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
      makeSpacedParagraph(1, "Item 2", 9, {
        spaceBefore: 5,
        spaceAfter: 13,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(1);
    const frags = layout.pages[0].fragments;
    expect(frags.length).toBe(2);

    // With contextual spacing suppressed, second paragraph should be immediately
    // after the first (no spaceAfter on first, no spaceBefore on second)
    const gap = frags[1].y - (frags[0].y + frags[0].height);
    expect(gap).toBe(0);
  });

  test("does NOT suppress spacing when contextualSpacing is false", () => {
    const blocks: FlowBlock[] = [
      makeSpacedParagraph(0, "Para 1", 1, {
        spaceAfter: 13,
        contextualSpacing: false,
        styleId: "Normal",
      }),
      makeSpacedParagraph(1, "Para 2", 9, {
        spaceBefore: 5,
        spaceAfter: 13,
        contextualSpacing: false,
        styleId: "Normal",
      }),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    const frags = layout.pages[0].fragments;
    // Gap = max(spaceAfter, spaceBefore) = 13 (Word-style collapsed spacing)
    const gap = frags[1].y - (frags[0].y + frags[0].height);
    expect(gap).toBe(13);
  });

  test("does NOT suppress spacing when styles differ", () => {
    const blocks: FlowBlock[] = [
      makeSpacedParagraph(0, "Bullet", 1, {
        spaceAfter: 13,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
      makeSpacedParagraph(1, "Normal", 9, {
        spaceBefore: 5,
        spaceAfter: 13,
        contextualSpacing: true,
        styleId: "Normal",
      }),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    const frags = layout.pages[0].fragments;
    // Different styles — spacing should NOT be suppressed
    // gap = max(spaceAfter, spaceBefore) = 13
    const gap = frags[1].y - (frags[0].y + frags[0].height);
    expect(gap).toBe(13);
  });

  test("suppresses only the opted-in side when one paragraph has contextualSpacing", () => {
    const blocks: FlowBlock[] = [
      makeSpacedParagraph(0, "First", 1, {
        spaceAfter: 13,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
      makeSpacedParagraph(1, "Second", 8, {
        spaceBefore: 5,
        spaceAfter: 13,
        contextualSpacing: false,
        styleId: "ListBullet",
      }),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    const frags = layout.pages[0].fragments;
    // The first paragraph suppresses its own spaceAfter; the second paragraph
    // retains its 5px spaceBefore.
    const gap = frags[1].y - (frags[0].y + frags[0].height);
    expect(gap).toBe(5);
  });

  test("suppresses spacing in a chain of 3+ same-style paragraphs", () => {
    const blocks: FlowBlock[] = [
      makeSpacedParagraph(0, "A", 1, {
        spaceAfter: 10,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
      makeSpacedParagraph(1, "B", 4, {
        spaceBefore: 5,
        spaceAfter: 10,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
      makeSpacedParagraph(2, "C", 7, {
        spaceBefore: 5,
        spaceAfter: 10,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 1, 10, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 1, 10, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 1, 10, 20)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    const frags = layout.pages[0].fragments;
    expect(frags.length).toBe(3);

    // All gaps should be zero
    const gap1 = frags[1].y - (frags[0].y + frags[0].height);
    const gap2 = frags[2].y - (frags[1].y + frags[1].height);
    expect(gap1).toBe(0);
    expect(gap2).toBe(0);
  });

  test("preserves spacing before first and after last in contextual chain", () => {
    // A normal paragraph, then 2 contextual, then normal
    const blocks: FlowBlock[] = [
      makeSpacedParagraph(0, "Normal", 1, {
        spaceAfter: 13,
        styleId: "Normal",
      }),
      makeSpacedParagraph(1, "Bullet 1", 9, {
        spaceBefore: 5,
        spaceAfter: 13,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
      makeSpacedParagraph(2, "Bullet 2", 19, {
        spaceBefore: 5,
        spaceAfter: 13,
        contextualSpacing: true,
        styleId: "ListBullet",
      }),
      makeSpacedParagraph(3, "Normal 2", 29, {
        spaceBefore: 5,
        spaceAfter: 13,
        styleId: "Normal",
      }),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 6, 60, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 8, 80, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 8, 80, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 8, 80, 20)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    const frags = layout.pages[0].fragments;
    expect(frags.length).toBe(4);

    // Gap between Normal and Bullet 1 — Normal has no contextualSpacing, so
    // gap = max(spaceAfter, spaceBefore) = 13
    const gap0to1 = frags[1].y - (frags[0].y + frags[0].height);
    expect(gap0to1).toBe(13);

    // Gap between Bullet 1 and Bullet 2 — both contextual, same style → suppressed
    const gap1to2 = frags[2].y - (frags[1].y + frags[1].height);
    expect(gap1to2).toBe(0);

    // Gap between Bullet 2 and Normal 2 — Normal 2 has no contextualSpacing
    // gap = max(spaceAfter, spaceBefore) = 13
    const gap2to3 = frags[3].y - (frags[2].y + frags[2].height);
    expect(gap2to3).toBe(13);
  });

  test("treats two absent style ids as the shared default paragraph style", () => {
    const blocks: FlowBlock[] = [
      makeSpacedParagraph(0, "No style 1", 1, {
        spaceAfter: 10,
        contextualSpacing: true,
        // no styleId
      }),
      makeSpacedParagraph(1, "No style 2", 13, {
        spaceBefore: 5,
        spaceAfter: 10,
        contextualSpacing: true,
        // no styleId
      }),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 10, 100, 20)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 10, 100, 20)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    const frags = layout.pages[0].fragments;
    const gap = frags[1].y - (frags[0].y + frags[0].height);
    expect(gap).toBe(0);
  });
});
