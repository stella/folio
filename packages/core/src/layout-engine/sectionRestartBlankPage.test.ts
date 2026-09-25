/**
 * Under `w:evenAndOddHeaders` (§17.10.1) page numbers alternate between odd
 * and even sheets. A `nextPage` section that restarts numbering with
 * `w:pgNumType w:start` (§17.6.12) with a number as odd or even as the sheet
 * before it gets a blank sheet in between.
 */

import { describe, expect, test } from "bun:test";

import { layoutDocument } from "./index";
import type {
  FlowBlock,
  Layout,
  Measure,
  PageHeaderFooterRefs,
  ParagraphBlock,
  ParagraphMeasure,
} from "./types";

const PAGE_SIZE = { w: 600, h: 800 };
const MARGINS = { top: 50, right: 50, bottom: 50, left: 50 };

const paragraph = (id: string): { block: ParagraphBlock; measure: ParagraphMeasure } => ({
  block: { kind: "paragraph", id, pmStart: 0, pmEnd: 0, runs: [{ kind: "text", text: id }] },
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
        lineHeight: 100,
      },
    ],
    totalHeight: 100,
  },
});

type Scenario = {
  evenAndOddHeaders: boolean;
  outgoingPages: number;
  restartAt: number;
};

const layoutRestartedSection = ({
  evenAndOddHeaders,
  outgoingPages,
  restartAt,
}: Scenario): Layout => {
  const refs: PageHeaderFooterRefs = {
    evenAndOddHeaders,
    headerDefault: "header-default",
    headerEven: "header-even",
    footerDefault: "footer-default",
  };
  const outgoing = Array.from({ length: outgoingPages }, (_, index) => paragraph(`out-${index}`));
  const incoming = paragraph("incoming");
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  outgoing.forEach(({ block, measure }, index) => {
    if (index > 0) {
      blocks.push({ kind: "pageBreak", id: `break-${index}` });
      measures.push({ kind: "pageBreak" });
    }
    blocks.push(block);
    measures.push(measure);
  });
  blocks.push(
    {
      kind: "sectionBreak",
      id: "boundary",
      pageSize: PAGE_SIZE,
      margins: MARGINS,
      pageNumbering: { type: "continue" },
    },
    incoming.block,
  );
  measures.push({ kind: "sectionBreak" }, incoming.measure);

  return layoutDocument(blocks, measures, {
    pageSize: PAGE_SIZE,
    margins: MARGINS,
    finalPageSize: PAGE_SIZE,
    finalMargins: MARGINS,
    finalPageNumbering: { type: "restart", start: restartAt },
    sectionHeaderFooterRefs: [refs, refs],
  });
};

const pageHolding = (layout: Layout, blockId: string) =>
  layout.pages.find((page) => page.fragments.some((fragment) => fragment.blockId === blockId));

describe("odd and even page numbers at a restarted section", () => {
  test("inserts a blank sheet when the restart repeats the previous page's odd or even number", () => {
    const layout = layoutRestartedSection({
      evenAndOddHeaders: true,
      outgoingPages: 1,
      restartAt: 1,
    });

    expect(layout.pages).toHaveLength(3);
    const filler = layout.pages[1]!;
    expect(filler.fragments).toHaveLength(0);
    // The filler names no header or footer part, so it paints no furniture.
    expect(filler.headerFooterRefs).toEqual({});
    const first = pageHolding(layout, "incoming");
    expect(first?.number).toBe(3);
    expect(first).toMatchObject({ logicalNumber: 1, sectionIndex: 1, sectionPageNumber: 1 });
    expect(first?.headerFooterRefs?.headerDefault).toBe("header-default");
  });

  test("continues on the next sheet when the restart alternates odd and even numbers", () => {
    const layout = layoutRestartedSection({
      evenAndOddHeaders: true,
      outgoingPages: 2,
      restartAt: 1,
    });

    expect(layout.pages).toHaveLength(3);
    expect(pageHolding(layout, "incoming")).toMatchObject({ number: 3, logicalNumber: 1 });
  });

  test("keeps consecutive sheets without even and odd headers", () => {
    const layout = layoutRestartedSection({
      evenAndOddHeaders: false,
      outgoingPages: 1,
      restartAt: 1,
    });

    expect(layout.pages).toHaveLength(2);
    expect(pageHolding(layout, "incoming")).toMatchObject({ number: 2, logicalNumber: 1 });
  });
});
