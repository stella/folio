import { describe, expect, test } from "bun:test";

import { getTableCellFloatingImages } from "./tableCellFloating";
import type {
  ImageRun,
  ParagraphBlock,
  ParagraphMeasure,
  TableCell,
  TableCellMeasure,
} from "../types";
import {
  createTableCellFlowState,
  finishTableCellFlow,
  placeTableCellBlock,
} from "./tableCellFlow";

const paragraph = (
  id: string,
  text: string,
  spacing: { before?: number; after?: number } = {},
): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: text.length > 0 ? [{ kind: "text", text }] : [],
  attrs: { spacing, styleId: "CellText" },
});

const measure = (contentHeight: number, spacing = 0): ParagraphMeasure => ({
  kind: "paragraph",
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 10,
      ascent: contentHeight * 0.8,
      descent: contentHeight * 0.2,
      lineHeight: contentHeight,
    },
  ],
  totalHeight: contentHeight + spacing,
});

const suppressedMeasure: ParagraphMeasure = {
  kind: "paragraph",
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 0,
      width: 0,
      ascent: 0,
      descent: 0,
      lineHeight: 0,
    },
  ],
  totalHeight: 0,
};

describe("table cell block flow", () => {
  test("collapses adjacent paragraph spacing to the larger side", () => {
    const state = createTableCellFlowState();
    placeTableCellBlock(state, paragraph("first", "first", { after: 8 }), measure(10, 8));
    const second = placeTableCellBlock(
      state,
      paragraph("second", "second", { before: 8 }),
      measure(10, 8),
    );

    expect(second.top).toBe(10);
    expect(second.leadingSpacing).toBe(8);
    expect(finishTableCellFlow(state)).toBe(28);
  });

  test("uses the paragraph box before leading spacing as its anchor origin", () => {
    const image: ImageRun = {
      kind: "image",
      src: "floating.png",
      width: 20,
      height: 20,
      displayMode: "float",
      wrapType: "square",
      position: { vertical: { relativeTo: "paragraph", posOffset: 0 } },
    };
    const cell: TableCell = {
      id: "cell",
      blocks: [
        paragraph("first", "first", { after: 8 }),
        {
          kind: "paragraph",
          id: "anchor",
          runs: [image],
          attrs: { spacing: { before: 12 } },
        },
      ],
    };
    const cellMeasure: TableCellMeasure = {
      blocks: [measure(10, 8), measure(10, 12)],
      width: 100,
      height: 32,
    };
    let pageResolverCalled = false;

    const images = getTableCellFloatingImages(cell, cellMeasure, 100, (_run, paragraphY) => {
      pageResolverCalled = true;
      return { side: "left", x: 0, y: paragraphY };
    });

    expect(pageResolverCalled).toBe(false);
    expect(images.at(0)?.y).toBe(10);
  });

  test("delegates anchors that explicitly escape the table cell", () => {
    const image: ImageRun = {
      kind: "image",
      src: "floating.png",
      width: 20,
      height: 20,
      displayMode: "float",
      wrapType: "square",
      layoutInCell: false,
      position: { vertical: { relativeTo: "paragraph", posOffset: 0 } },
    };
    const cell: TableCell = {
      id: "cell",
      blocks: [{ kind: "paragraph", id: "anchor", runs: [image] }],
    };
    const cellMeasure: TableCellMeasure = {
      blocks: [measure(10)],
      width: 100,
      height: 10,
    };

    const images = getTableCellFloatingImages(cell, cellMeasure, 100, (_run, paragraphY) => ({
      side: "right",
      x: -40,
      y: paragraphY + 5,
    }));

    expect(images.at(0)).toMatchObject({ x: -40, y: 5, side: "right" });
  });

  test("resolves default table-cell anchor offsets from the cell content origin", () => {
    const image: ImageRun = {
      kind: "image",
      src: "floating.png",
      width: 20,
      height: 20,
      displayMode: "float",
      wrapType: "square",
      position: {
        horizontal: { relativeTo: "margin", posOffset: -914_400 },
        vertical: { relativeTo: "paragraph", posOffset: 0 },
      },
    };
    const cell: TableCell = {
      id: "cell",
      blocks: [{ kind: "paragraph", id: "anchor", runs: [image] }],
    };
    const cellMeasure: TableCellMeasure = {
      blocks: [measure(10)],
      width: 100,
      height: 10,
    };
    let pageResolverCalled = false;

    const images = getTableCellFloatingImages(cell, cellMeasure, 100, () => {
      pageResolverCalled = true;
      return { side: "left", x: -96, y: 0 };
    });

    expect(pageResolverCalled).toBe(false);
    expect(images.at(0)).toMatchObject({ x: -96, y: 0, side: "left" });
  });

  test("keeps consecutive authored empty paragraphs as independent spacers", () => {
    const state = createTableCellFlowState();
    placeTableCellBlock(state, paragraph("body", "body", { after: 8 }), measure(10, 8));
    placeTableCellBlock(state, paragraph("blank-1", "", { before: 8, after: 8 }), measure(10, 16));
    const secondBlank = placeTableCellBlock(
      state,
      paragraph("blank-2", "", { before: 8, after: 8 }),
      measure(10, 16),
    );

    expect(secondBlank.leadingSpacing).toBe(16);
    expect(finishTableCellFlow(state)).toBe(62);
  });

  test("includes final authored paragraph spacing in the cell height", () => {
    const state = createTableCellFlowState();
    placeTableCellBlock(state, paragraph("body", "body", { after: 8 }), measure(10, 8));

    expect(state.height).toBe(10);
    expect(finishTableCellFlow(state)).toBe(18);
  });

  test("keeps a suppressed empty paragraph at zero height", () => {
    const state = createTableCellFlowState();
    const block = paragraph("marker", "", { before: 8, after: 8 });
    block.attrs = { ...block.attrs, suppressEmptyParagraphHeight: true };

    expect(placeTableCellBlock(state, block, suppressedMeasure).leadingSpacing).toBe(0);
    expect(finishTableCellFlow(state)).toBe(0);
  });

  test("collapses surrounding spacing across a suppressed empty paragraph", () => {
    const state = createTableCellFlowState();
    placeTableCellBlock(state, paragraph("first", "first", { after: 8 }), measure(10, 8));
    const marker = paragraph("marker", "", { before: 8, after: 8 });
    marker.attrs = { ...marker.attrs, suppressEmptyParagraphHeight: true };
    const markerPlacement = placeTableCellBlock(state, marker, suppressedMeasure);
    const third = placeTableCellBlock(
      state,
      paragraph("third", "third", { before: 8 }),
      measure(10, 8),
    );

    expect(markerPlacement).toEqual({
      top: 10,
      contentTop: 10,
      contentHeight: 0,
      leadingSpacing: 0,
    });
    expect(third.leadingSpacing).toBe(8);
    expect(finishTableCellFlow(state)).toBe(28);
  });
});
