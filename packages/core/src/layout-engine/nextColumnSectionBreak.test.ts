/**
 * `nextColumn` is the fifth `ST_SectionMark` member (ECMA-376 Part 1
 * §17.18.77). It starts the incoming section in the next column of the column
 * region the outgoing section occupies.
 *
 * The member list comes from the committed schema graph rather than a hand
 * list here, so a schema refresh that adds a member fails this file instead of
 * reaching the paginator as an unmodelled value.
 */

import { readFileSync } from "node:fs";

import { describe, expect, test } from "bun:test";

import type { SectionStart } from "@stll/docx-core/model";

import { layoutDocument } from "./index";
import { calculateColumnLefts, calculateColumnWidths } from "./paginator";
import { normalizeSectionBreakType } from "./section-breaks";
import type { FlowBlock, Measure, Page, ParagraphBlock, ParagraphMeasure } from "./types";

const ST_SECTION_MARK_VALUES: readonly string[] = (() => {
  const graph = JSON.parse(
    readFileSync(
      new URL(
        "../../../../specifications/generated/docx-transitional-schema.gen.json",
        import.meta.url,
      ),
      "utf8",
    ),
  ) as { symbols: { kind?: string; name?: string; enumValues?: string[] }[] };
  const sectionMark = graph.symbols.find(
    (symbol) => symbol.kind === "simpleType" && symbol.name === "ST_SectionMark",
  );
  if (!sectionMark?.enumValues) {
    throw new Error("ST_SectionMark is missing from the schema graph");
  }
  return sectionMark.enumValues;
})();

const PAGE_SIZE = { w: 800, h: 1000 };
const MARGINS = { top: 50, right: 50, bottom: 50, left: 50 };
const TWO_COLUMNS = { count: 2, gap: 20 };
const PARAGRAPH_HEIGHT = 100;

const paragraph = (id: string): { block: ParagraphBlock; measure: ParagraphMeasure } => ({
  block: {
    kind: "paragraph",
    id,
    pmStart: 0,
    pmEnd: 0,
    runs: [{ kind: "text", text: id }],
    attrs: {},
  },
  measure: {
    kind: "paragraph",
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 0,
        width: 100,
        ascent: 10,
        descent: 3,
        lineHeight: PARAGRAPH_HEIGHT,
      },
    ],
    totalHeight: PARAGRAPH_HEIGHT,
  },
});

/**
 * The column a fragment sits in, read back from the page's own column model
 * rather than from a pixel constant: the assertion then states placement, not
 * arithmetic.
 */
const columnIndexOf = (page: Page, blockId: string): number => {
  const fragment = page.fragments.find(
    (candidate) => candidate.kind === "paragraph" && candidate.blockId === blockId,
  );
  if (!fragment) {
    throw new Error(`no fragment for ${blockId}`);
  }
  const columns = page.columns ?? { count: 1, gap: 0 };
  const lefts = calculateColumnLefts(
    page.margins.left,
    calculateColumnWidths(page.size.w, page.margins.left, page.margins.right, columns),
    columns,
  );
  const index = lefts.findIndex((left) => Math.round(left) === Math.round(fragment.x));
  if (index === -1) {
    throw new Error(`${blockId} at x=${fragment.x} matches no column of ${JSON.stringify(lefts)}`);
  }
  return index;
};

/**
 * Where the incoming section's first block lands: 1-based page, 0-based
 * column, and the section index that page is recorded against. A boundary that
 * never opens a section leaves the outgoing section's index on the page.
 */
type Start = { page: number; column: number; sectionIndex: number | undefined };

const startOfIncomingSection = (
  breakType: SectionStart | undefined,
  leadingParagraphs: readonly string[],
): Start => {
  const leading = leadingParagraphs.map((id) => paragraph(id));
  const incoming = paragraph("incoming");
  const blocks: FlowBlock[] = [
    ...leading.map(({ block }) => block),
    {
      kind: "sectionBreak",
      id: "boundary",
      pageSize: PAGE_SIZE,
      margins: MARGINS,
      columns: TWO_COLUMNS,
      ...(breakType === undefined ? {} : { type: breakType }),
    },
    incoming.block,
  ];
  const measures: Measure[] = [
    ...leading.map(({ measure }) => measure),
    { kind: "sectionBreak" },
    incoming.measure,
  ];

  const layout = layoutDocument(blocks, measures, {
    pageSize: PAGE_SIZE,
    margins: MARGINS,
    finalPageSize: PAGE_SIZE,
    finalMargins: MARGINS,
    finalColumns: TWO_COLUMNS,
  });

  const pageIndex = layout.pages.findIndex((page) =>
    page.fragments.some(
      (fragment) => fragment.kind === "paragraph" && fragment.blockId === "incoming",
    ),
  );
  const page = layout.pages[pageIndex];
  if (!page) {
    throw new Error("the incoming section was not laid out");
  }
  return {
    page: pageIndex + 1,
    column: columnIndexOf(page, "incoming"),
    sectionIndex: page.sectionIndex,
  };
};

/**
 * Where each member starts the incoming section, given a two-column outgoing
 * section whose first column already holds one paragraph. Total over the
 * enumeration, so a new member needs a decision here rather than inheriting one.
 */
const INCOMING_SECTION_START = [
  { type: "nextPage", start: { page: 2, column: 0, sectionIndex: 1 } },
  // The next column of the sheet the outgoing section still owns.
  { type: "nextColumn", start: { page: 1, column: 1, sectionIndex: 0 } },
  { type: "continuous", start: { page: 1, column: 0, sectionIndex: 0 } },
  { type: "evenPage", start: { page: 2, column: 0, sectionIndex: 1 } },
  { type: "oddPage", start: { page: 3, column: 0, sectionIndex: 1 } },
] as const satisfies readonly { type: SectionStart; start: Start }[];

describe("ST_SectionMark coverage", () => {
  test("the layout decides every member the schema enumerates", () => {
    expect(INCOMING_SECTION_START.map(({ type }) => type).toSorted()).toEqual(
      [...ST_SECTION_MARK_VALUES].sort(),
    );
  });

  test.each(INCOMING_SECTION_START)(
    "$type starts the incoming section where Word does",
    ({ type, start }) => {
      expect(startOfIncomingSection(type, ["outgoing"])).toEqual(start);
    },
  );

  test("an absent w:type starts the next page (§17.6.22)", () => {
    expect(normalizeSectionBreakType(undefined)).toBe("nextPage");
    expect(startOfIncomingSection(undefined, ["outgoing"])).toEqual({
      page: 2,
      column: 0,
      sectionIndex: 1,
    });
  });

  test("a value outside the enumeration is a bug, not a default", () => {
    // SAFETY: the point of the assertion is a value the type forbids.
    expect(() => normalizeSectionBreakType("nextFrame" as SectionStart)).toThrow();
  });
});

describe("nextColumn", () => {
  test("a break in the last column starts the incoming section on a new page", () => {
    // 900px of body per column holds nine paragraphs, so the tenth already
    // sits in the last column and the break has no further column to use.
    const outgoing = Array.from({ length: 10 }, (_, index) => `outgoing-${index}`);

    // The sheet it opens belongs to the incoming section: the boundary starts
    // one, so the section's page numbering and furniture take effect there.
    expect(startOfIncomingSection("nextColumn", outgoing)).toEqual({
      page: 2,
      column: 0,
      sectionIndex: 1,
    });
  });

  test("a single-column section has no next column, so the break falls back", () => {
    const outgoing = paragraph("outgoing");
    const incoming = paragraph("incoming");
    const blocks: FlowBlock[] = [
      outgoing.block,
      { kind: "sectionBreak", id: "boundary", type: "nextColumn" },
      incoming.block,
    ];
    const measures: Measure[] = [outgoing.measure, { kind: "sectionBreak" }, incoming.measure];

    const layout = layoutDocument(blocks, measures, { pageSize: PAGE_SIZE, margins: MARGINS });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]?.fragments.map((fragment) => fragment.blockId)).toEqual([
      "outgoing",
      "incoming",
    ]);
    expect(layout.pages[0]?.fragments.map((fragment) => fragment.y)).toEqual([
      MARGINS.top,
      MARGINS.top + PARAGRAPH_HEIGHT,
    ]);
  });

  test("a changed physical region falls back in place instead of borrowing a column", () => {
    const outgoing = paragraph("outgoing");
    const incoming = paragraph("incoming");
    const blocks: FlowBlock[] = [
      outgoing.block,
      {
        kind: "sectionBreak",
        id: "boundary",
        type: "nextColumn",
        pageSize: PAGE_SIZE,
        margins: MARGINS,
        columns: TWO_COLUMNS,
      },
      incoming.block,
    ];
    const measures: Measure[] = [outgoing.measure, { kind: "sectionBreak" }, incoming.measure];
    const incomingMargins = { ...MARGINS, left: 100, right: 100 };

    const layout = layoutDocument(blocks, measures, {
      pageSize: PAGE_SIZE,
      margins: MARGINS,
      columns: TWO_COLUMNS,
      finalPageSize: PAGE_SIZE,
      finalMargins: incomingMargins,
      finalColumns: TWO_COLUMNS,
    });
    const page = layout.pages.at(0);

    expect(layout.pages).toHaveLength(1);
    expect(page?.margins).toEqual(MARGINS);
    expect(page ? columnIndexOf(page, "incoming") : undefined).toBe(0);
    expect(page?.fragments.find((fragment) => fragment.blockId === "incoming")?.y).toBe(
      MARGINS.top + PARAGRAPH_HEIGHT,
    );
  });

  test("a changed page size starts nextColumn on a new sheet", () => {
    const outgoing = paragraph("outgoing");
    const incoming = paragraph("incoming");
    const blocks: FlowBlock[] = [
      outgoing.block,
      {
        kind: "sectionBreak",
        id: "boundary",
        type: "nextColumn",
        pageSize: PAGE_SIZE,
        margins: MARGINS,
        columns: TWO_COLUMNS,
      },
      incoming.block,
    ];
    const measures: Measure[] = [outgoing.measure, { kind: "sectionBreak" }, incoming.measure];
    const incomingPageSize = { w: 1000, h: 800 };

    const layout = layoutDocument(blocks, measures, {
      pageSize: PAGE_SIZE,
      margins: MARGINS,
      columns: TWO_COLUMNS,
      finalPageSize: incomingPageSize,
      finalMargins: MARGINS,
      finalColumns: TWO_COLUMNS,
    });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.at(1)?.size).toEqual(incomingPageSize);
    expect(layout.pages.at(1)?.fragments.map((fragment) => fragment.blockId)).toEqual(["incoming"]);
  });
});
