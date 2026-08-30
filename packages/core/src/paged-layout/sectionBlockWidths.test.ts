import { describe, expect, test } from "bun:test";

import type { SectionLayoutConfig } from "../layout-engine";
import type { FlowBlock, ParagraphBlock } from "../layout-engine/types";
import { computePerBlockMeasureInputs, computePerBlockWidths } from "./sectionBlockWidths";

const BODY_CONFIG = {
  pageSize: { w: 1000, h: 1200 },
  margins: { top: 50, right: 100, bottom: 50, left: 100 },
  pageNumbering: { type: "continue" },
} satisfies SectionLayoutConfig;

function paragraph(id: string): ParagraphBlock {
  return {
    kind: "paragraph",
    id,
    runs: [{ kind: "text", text: id }],
    attrs: {},
  };
}

describe("section block measurement inputs", () => {
  test("uses each section's page size and margins for measurement width", () => {
    const blocks: FlowBlock[] = [
      paragraph("first-section"),
      {
        kind: "sectionBreak",
        id: "section-one",
        pageSize: { w: 600, h: 1200 },
        margins: { top: 50, right: 50, bottom: 50, left: 50 },
      },
      paragraph("final-section"),
    ];

    const widths = computePerBlockWidths({
      blocks,
      bodyConfig: BODY_CONFIG,
      finalConfig: BODY_CONFIG,
    });

    expect(widths).toEqual([500, 500, 800]);
  });

  test("returns each block's section top margin for band measurement", () => {
    const blocks: FlowBlock[] = [
      paragraph("first-section"),
      {
        kind: "sectionBreak",
        id: "section-one",
        pageSize: { w: 600, h: 1200 },
        margins: { top: 64, right: 50, bottom: 50, left: 50 },
      },
      paragraph("final-section"),
    ];

    const inputs = computePerBlockMeasureInputs({
      blocks,
      bodyConfig: BODY_CONFIG,
      finalConfig: BODY_CONFIG,
    });

    expect(inputs.marginTops).toEqual([64, 64, BODY_CONFIG.margins.top]);
    expect(inputs.contentTops).toEqual([64, 64, BODY_CONFIG.margins.top]);
    expect(inputs.pageWidths).toEqual([600, 600, BODY_CONFIG.pageSize.w]);
    expect(inputs.marginLefts).toEqual([50, 50, BODY_CONFIG.margins.left]);
    expect(inputs.marginRights).toEqual([50, 50, BODY_CONFIG.margins.right]);
    expect(inputs.contentLefts).toEqual([50, 50, BODY_CONFIG.margins.left]);
    expect(inputs.columnIndices).toEqual([0, 0, 0]);
    expect(inputs.columnCounts).toEqual([1, 1, 1]);
  });

  test("uses the authored width after a forced break in unequal columns", () => {
    const blocks: FlowBlock[] = [
      paragraph("first-column"),
      { kind: "columnBreak", id: "column-break" },
      paragraph("second-column"),
    ];

    const inputs = computePerBlockMeasureInputs({
      blocks,
      bodyConfig: BODY_CONFIG,
      finalConfig: {
        ...BODY_CONFIG,
        columns: { count: 2, gap: 20, widths: [200, 300], gaps: [100] },
      },
    });

    expect(inputs.widths).toEqual([200, 200, 300]);
    expect(inputs.contentLefts).toEqual([100, 100, 400]);
    expect(inputs.columnIndices).toEqual([0, 0, 1]);
    expect(inputs.columnCounts).toEqual([2, 2, 2]);
  });

  test("tracks active origins across unequal columns and a page reset", () => {
    const blocks: FlowBlock[] = [
      paragraph("first"),
      { kind: "columnBreak", id: "first-break" },
      paragraph("second"),
      { kind: "columnBreak", id: "second-break" },
      paragraph("third"),
      { kind: "pageBreak", id: "page-break" },
      paragraph("next-page"),
    ];
    const inputs = computePerBlockMeasureInputs({
      blocks,
      bodyConfig: BODY_CONFIG,
      finalConfig: {
        ...BODY_CONFIG,
        columns: { count: 3, gap: 0, widths: [180, 260, 220], gaps: [30, 70] },
      },
    });

    expect(inputs.widths).toEqual([180, 180, 260, 260, 220, 220, 180]);
    expect(inputs.contentLefts).toEqual([100, 100, 310, 310, 640, 640, 100]);
  });

  test("keeps the outgoing physical content top until a shared continuous page advances", () => {
    const nextSection = {
      ...BODY_CONFIG,
      margins: { ...BODY_CONFIG.margins, top: 200 },
    } satisfies SectionLayoutConfig;
    const blocks: FlowBlock[] = [
      paragraph("outgoing"),
      {
        kind: "sectionBreak",
        id: "continuous",
        type: "continuous",
        pageSize: BODY_CONFIG.pageSize,
        margins: { ...BODY_CONFIG.margins, top: 64 },
      },
      paragraph("shared-page"),
      { kind: "pageBreak", id: "new-page" },
      paragraph("new-page-body"),
    ];

    const inputs = computePerBlockMeasureInputs({
      blocks,
      bodyConfig: BODY_CONFIG,
      finalConfig: nextSection,
    });

    expect(inputs.marginTops).toEqual([64, 64, 200, 200, 200]);
    expect(inputs.contentTops).toEqual([64, 64, 64, 64, 200]);
  });
});
