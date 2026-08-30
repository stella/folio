import { describe, expect, test } from "bun:test";

import type { SectionLayoutConfig } from "../layout-engine";
import { layoutDocument } from "../layout-engine";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../layout-engine/measure/measureBlocks";
import type { FlowBlock, ParagraphBlock, TableBlock } from "../layout-engine/types";
import { computePerBlockMeasureInputs, computePerBlockWidths } from "./sectionBlockWidths";

const BODY_CONFIG = {
  pageSize: { w: 1000, h: 1200 },
  margins: { top: 50, right: 100, bottom: 50, left: 100 },
  pageNumbering: { type: "continue" },
} satisfies SectionLayoutConfig;

const fakeMeasure = { charWidth: fixedCharWidth(5) };

function paragraph(id: string): ParagraphBlock {
  return {
    kind: "paragraph",
    id,
    runs: [{ kind: "text", text: id }],
    attrs: {},
  };
}

function centeredMarginFloatingTable(id: string): TableBlock {
  return {
    kind: "table",
    id,
    columnWidths: [200],
    floating: {
      horzAnchor: "margin",
      vertAnchor: "text",
      tblpXSpec: "center",
      tblpY: 0,
    },
    rows: [
      {
        id: `${id}-row`,
        cells: [
          {
            id: `${id}-cell`,
            blocks: [paragraph(`${id}-content`)],
          },
        ],
      },
    ],
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
    expect(inputs.physicalPageGeometry.marginTops).toEqual([64, 64, BODY_CONFIG.margins.top]);
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

  test("resets an authored column before a pageBreakBefore paragraph", () => {
    const pageBreakParagraph = paragraph("new-page-first-column");
    pageBreakParagraph.attrs = { pageBreakBefore: true };
    const blocks: FlowBlock[] = [
      paragraph("first-column"),
      { kind: "columnBreak", id: "column-break" },
      paragraph("second-column"),
      pageBreakParagraph,
    ];

    const inputs = computePerBlockMeasureInputs({
      blocks,
      bodyConfig: BODY_CONFIG,
      finalConfig: {
        ...BODY_CONFIG,
        columns: { count: 2, gap: 20 },
      },
    });

    expect(inputs.columnIndices).toEqual([0, 0, 1, 0]);
    expect(inputs.contentLefts).toEqual([100, 100, 510, 100]);
    expect(inputs.widths).toEqual([390, 390, 390, 390]);
  });

  test("keeps the complete outgoing physical geometry until a shared continuous page advances", () => {
    const nextSection = {
      ...BODY_CONFIG,
      margins: { top: 200, right: 140, bottom: 180, left: 160 },
      columns: { count: 2, gap: 20 },
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
    expect(inputs.physicalPageGeometry).toEqual({
      pageHeights: [1200, 1200, 1200, 1200, 1200],
      pageWidths: [1000, 1000, 1000, 1000, 1000],
      marginTops: [64, 64, 64, 64, 200],
      marginLefts: [100, 100, 100, 100, 160],
      marginRights: [100, 100, 100, 100, 140],
      marginBottoms: [50, 50, 50, 50, 180],
    });
    expect(inputs.widths).toEqual([800, 800, 390, 390, 340]);
    expect(inputs.contentLefts).toEqual([100, 100, 100, 100, 160]);
    expect(inputs.columnCounts).toEqual([1, 1, 2, 2, 2]);
  });

  test("aligns a new margin table's measured exclusion with its retained-page fragment", () => {
    withFakeTextMeasure(() => {
      const outgoingConfig = {
        pageSize: { w: 800, h: 1_000 },
        margins: { top: 40, right: 160, bottom: 40, left: 40 },
        pageNumbering: { type: "continue" },
      } as const;
      const incomingConfig = {
        pageSize: { w: 800, h: 1_000 },
        margins: { top: 200, right: 100, bottom: 180, left: 200 },
        pageNumbering: { type: "continue" },
      } as const;
      const table = centeredMarginFloatingTable("incoming-margin-float");
      const blocks: FlowBlock[] = [
        paragraph("outgoing-body"),
        {
          kind: "sectionBreak",
          id: "continuous-margin-change",
          type: "continuous",
          pageSize: outgoingConfig.pageSize,
          margins: outgoingConfig.margins,
        },
        table,
        paragraph("incoming-body"),
      ];
      const inputs = computePerBlockMeasureInputs({
        blocks,
        bodyConfig: outgoingConfig,
        finalConfig: incomingConfig,
      });
      const measures = measureBlocks(blocks, inputs.widths, inputs.marginTops, {
        pageWidth: inputs.pageWidths,
        pageHeight: inputs.pageHeights,
        marginLeft: inputs.marginLefts,
        marginRight: inputs.marginRights,
        marginBottom: inputs.marginBottoms,
        contentLeft: inputs.contentLefts,
        physicalPage: {
          pageWidth: inputs.physicalPageGeometry.pageWidths,
          pageHeight: inputs.physicalPageGeometry.pageHeights,
          marginTop: inputs.physicalPageGeometry.marginTops,
          marginLeft: inputs.physicalPageGeometry.marginLefts,
          marginRight: inputs.physicalPageGeometry.marginRights,
          marginBottom: inputs.physicalPageGeometry.marginBottoms,
        },
        columnIndex: inputs.columnIndices,
        columnCount: inputs.columnCounts,
      });
      const layout = layoutDocument(blocks, measures, {
        pageSize: outgoingConfig.pageSize,
        margins: outgoingConfig.margins,
        finalPageSize: incomingConfig.pageSize,
        finalMargins: incomingConfig.margins,
      });
      const bodyMeasure = measures.at(3);
      const tableMeasure = measures.at(2);
      const tableFragment = layout.pages
        .flatMap((page) => page.fragments)
        .find((fragment) => fragment.kind === "table" && fragment.blockId === table.id);
      if (bodyMeasure?.kind !== "paragraph" || tableMeasure?.kind !== "table") {
        throw new Error("Expected paragraph and table measures");
      }

      expect(tableFragment?.x).toBe(240);
      expect(bodyMeasure.lines.at(0)?.leftOffset).toBe(
        (tableFragment?.x ?? 0) + tableMeasure.totalWidth + 12 - (inputs.contentLefts.at(3) ?? 0),
      );
    }, fakeMeasure);
  });

  test("adopts pending geometry and clears stale zones before pageBreakBefore content", () => {
    withFakeTextMeasure(() => {
      const outgoingConfig = {
        pageSize: { w: 800, h: 1_000 },
        margins: { top: 40, right: 160, bottom: 40, left: 40 },
        pageNumbering: { type: "continue" },
      } as const;
      const incomingConfig = {
        pageSize: { w: 800, h: 1_000 },
        margins: { top: 200, right: 100, bottom: 180, left: 200 },
        pageNumbering: { type: "continue" },
      } as const;
      const outgoingTable = centeredMarginFloatingTable("outgoing-margin-float");
      const incomingTable = centeredMarginFloatingTable("incoming-margin-float");
      const pageBreakParagraph = paragraph("page-break-before");
      pageBreakParagraph.attrs = { pageBreakBefore: true };
      const blocks: FlowBlock[] = [
        outgoingTable,
        paragraph("outgoing-body"),
        {
          kind: "sectionBreak",
          id: "continuous-margin-change",
          type: "continuous",
          pageSize: outgoingConfig.pageSize,
          margins: outgoingConfig.margins,
        },
        pageBreakParagraph,
        incomingTable,
        paragraph("incoming-body"),
      ];
      const inputs = computePerBlockMeasureInputs({
        blocks,
        bodyConfig: outgoingConfig,
        finalConfig: incomingConfig,
      });
      const measures = measureBlocks(blocks, inputs.widths, inputs.marginTops, {
        pageWidth: inputs.pageWidths,
        pageHeight: inputs.pageHeights,
        marginLeft: inputs.marginLefts,
        marginRight: inputs.marginRights,
        marginBottom: inputs.marginBottoms,
        contentLeft: inputs.contentLefts,
        physicalPage: {
          pageWidth: inputs.physicalPageGeometry.pageWidths,
          pageHeight: inputs.physicalPageGeometry.pageHeights,
          marginTop: inputs.physicalPageGeometry.marginTops,
          marginLeft: inputs.physicalPageGeometry.marginLefts,
          marginRight: inputs.physicalPageGeometry.marginRights,
          marginBottom: inputs.physicalPageGeometry.marginBottoms,
        },
        columnIndex: inputs.columnIndices,
        columnCount: inputs.columnCounts,
      });
      const layout = layoutDocument(blocks, measures, {
        pageSize: outgoingConfig.pageSize,
        margins: outgoingConfig.margins,
        finalPageSize: incomingConfig.pageSize,
        finalMargins: incomingConfig.margins,
      });
      const pageBreakMeasure = measures.at(3);
      const tableMeasure = measures.at(4);
      const bodyMeasure = measures.at(5);
      const tableFragment = layout.pages
        .flatMap((page) => page.fragments)
        .find((fragment) => fragment.kind === "table" && fragment.blockId === incomingTable.id);
      if (
        pageBreakMeasure?.kind !== "paragraph" ||
        tableMeasure?.kind !== "table" ||
        bodyMeasure?.kind !== "paragraph"
      ) {
        throw new Error("Expected paragraph and table measures");
      }

      expect(inputs.physicalPageGeometry.marginLefts).toEqual([40, 40, 40, 200, 200, 200]);
      expect(inputs.physicalPageGeometry.marginRights).toEqual([160, 160, 160, 100, 100, 100]);
      expect(inputs.widths).toEqual([600, 600, 600, 500, 500, 500]);
      expect(inputs.contentLefts).toEqual([40, 40, 40, 200, 200, 200]);
      expect(pageBreakMeasure.lines.at(0)?.leftOffset).toBeUndefined();
      expect(pageBreakMeasure.lines.at(0)?.rightOffset).toBeUndefined();
      expect(tableFragment?.x).toBe(350);
      expect(bodyMeasure.lines.at(0)?.leftOffset).toBe(
        (tableFragment?.x ?? 0) + tableMeasure.totalWidth + 12 - (inputs.contentLefts.at(5) ?? 0),
      );
    }, fakeMeasure);
  });
});
