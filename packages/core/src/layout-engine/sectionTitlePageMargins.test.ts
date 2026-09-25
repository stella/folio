/**
 * `w:titlePg` (§17.10.6) is a property of each section: the first page of a
 * section that sets it shows the section's first-page header and footer, so its
 * body clears those parts, while the first page of a section that does not set
 * it keeps the margins of its default parts.
 */

import { describe, expect, test } from "bun:test";

import { layoutDocument } from "./index";
import type { FlowBlock, Measure, PageMargins, ParagraphBlock, ParagraphMeasure } from "./types";

const PAGE_SIZE = { w: 600, h: 800 };
const MARGINS: PageMargins = { top: 50, right: 50, bottom: 50, left: 50 };
const TITLE_PAGE_MARGINS: PageMargins = { ...MARGINS, top: 120 };

const paragraph = (id: string): [ParagraphBlock, ParagraphMeasure] => [
  { kind: "paragraph", id, pmStart: 0, pmEnd: 0, runs: [{ kind: "text", text: id }] },
  {
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
        lineHeight: 100,
      },
    ],
    totalHeight: 100,
  },
];

/** Two sections of two pages each, split by hard page breaks. */
const layoutTwoSections = (sectionFirstPageMargins: (PageMargins | undefined)[]) => {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  const push = (block: FlowBlock, measure: Measure) => {
    blocks.push(block);
    measures.push(measure);
  };
  const pageBreak = (id: string) => push({ kind: "pageBreak", id }, { kind: "pageBreak" });
  push(...paragraph("s0-p1"));
  pageBreak("s0-break");
  push(...paragraph("s0-p2"));
  push(
    { kind: "sectionBreak", id: "boundary", pageSize: PAGE_SIZE, margins: MARGINS },
    { kind: "sectionBreak" },
  );
  push(...paragraph("s1-p1"));
  pageBreak("s1-break");
  push(...paragraph("s1-p2"));

  return layoutDocument(blocks, measures, {
    pageSize: PAGE_SIZE,
    margins: MARGINS,
    finalPageSize: PAGE_SIZE,
    finalMargins: MARGINS,
    sectionFirstPageMargins,
  });
};

describe("title page margins per section", () => {
  test("applies only to the first page of the section that sets w:titlePg", () => {
    const layout = layoutTwoSections([undefined, TITLE_PAGE_MARGINS]);

    expect(layout.pages.map((page) => page.margins.top)).toEqual([50, 50, 120, 50]);
    expect(layout.pages[2]?.fragments.at(0)).toMatchObject({ blockId: "s1-p1", y: 120 });
  });

  test("leaves every page on the section margins without title pages", () => {
    const layout = layoutTwoSections([undefined, undefined]);

    expect(layout.pages.map((page) => page.margins.top)).toEqual([50, 50, 50, 50]);
  });
});
