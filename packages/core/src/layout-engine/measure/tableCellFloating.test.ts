/**
 * Wrap exclusions for pictures anchored inside a table cell. A picture that
 * leaves no room beside it (`wp:wrapSquare` spanning the cell, or
 * `wp:wrapTopAndBottom`) must push the cell's text below it; `wp:wrapNone`
 * pictures (`behindDoc` / in front of text) exclude nothing.
 */

import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { clearAllCaches } from "./cache";
import { measureTableBlock } from "./measureBlocks";
import { installCanvasMeasureProvider, resetCanvasContext } from "./measureContainer";
import { measureParagraph } from "./measureParagraph";
import {
  buildTableCellFloatingZones,
  getTableCellContentWidth,
  getTableCellFloatingImages,
} from "./tableCellFloating";
import type { ImageRun, ParagraphBlock, ParagraphMeasure, TableBlock } from "../types";

const CELL_WIDTH = 200;
const PICTURE_HEIGHT = 80;
const originalDocument = globalThis.document;

beforeEach(() => {
  installCanvasMeasureProvider();
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: {
      createElement(tagName: string) {
        if (tagName !== "canvas") {
          return {};
        }
        return {
          getContext() {
            return {
              font: "",
              measureText(text: string) {
                return { width: text.length * 7 };
              },
            };
          },
        };
      },
    },
  });
  clearAllCaches();
  resetCanvasContext();
});

afterEach(() => {
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: originalDocument,
  });
  clearAllCaches();
  resetCanvasContext();
});

function picture(overrides: Partial<ImageRun>): ImageRun {
  return {
    kind: "image",
    src: "picture.png",
    width: CELL_WIDTH,
    height: PICTURE_HEIGHT,
    wrapType: "square",
    displayMode: "float",
    distTop: 0,
    distBottom: 0,
    distLeft: 0,
    distRight: 0,
    position: {
      horizontal: { relativeTo: "column", posOffset: 0 },
      vertical: { relativeTo: "paragraph", posOffset: 0 },
    },
    ...overrides,
  };
}

function cellTable(run: ImageRun): TableBlock {
  const paragraph: ParagraphBlock = {
    kind: "paragraph",
    id: "p0",
    runs: [run, { kind: "text", text: "caption text under the picture" }],
  };
  return {
    kind: "table",
    id: "t",
    rows: [
      {
        id: "r0",
        cells: [
          {
            id: "c0",
            padding: { top: 0, right: 0, bottom: 0, left: 0 },
            blocks: [paragraph],
          },
        ],
      },
    ],
    columnWidths: [CELL_WIDTH],
  };
}

function laidOutCell(run: ImageRun) {
  const block = cellTable(run);
  const measure = measureTableBlock(block, CELL_WIDTH);
  const cell = block.rows[0]!.cells[0]!;
  const cellMeasure = measure.rows[0]!.cells[0]!;
  const contentWidth = getTableCellContentWidth(cell, cellMeasure);
  const zones = buildTableCellFloatingZones(
    getTableCellFloatingImages(cell, cellMeasure, contentWidth),
    contentWidth,
  );
  const paragraph = cell.blocks[0] as ParagraphBlock;
  const wrapped = measureParagraph(paragraph, contentWidth, { floatingZones: zones });
  return { measure, cellMeasure, zones, wrapped };
}

/** Y of the first text line, below any float skip. */
const firstLineTop = (measure: ParagraphMeasure): number => measure.lines[0]?.floatSkipBefore ?? 0;

describe("table cell floating picture exclusions", () => {
  test.each<[string, Partial<ImageRun>]>([
    ["square, exactly the cell width", {}],
    ["square, wider than the cell", { width: CELL_WIDTH + 40 }],
    ["square, its wrap distance reaching the cell edge", { width: CELL_WIDTH - 6, distRight: 12 }],
    ["topAndBottom", { wrapType: "topAndBottom" }],
  ])("%s pushes the cell text below the picture", (_label, overrides) => {
    const { zones, wrapped } = laidOutCell(picture(overrides));

    expect(zones).toHaveLength(1);
    expect(zones[0]?.fullWidthBlock).toBe(true);
    expect(firstLineTop(wrapped)).toBeGreaterThanOrEqual(PICTURE_HEIGHT);
  });

  test("the row grows to hold the text pushed below the picture", () => {
    const { measure, cellMeasure } = laidOutCell(picture({}));
    const stored = cellMeasure.blocks[0] as ParagraphMeasure;

    expect(firstLineTop(stored)).toBeGreaterThanOrEqual(PICTURE_HEIGHT);
    expect(measure.rows[0]!.height).toBeGreaterThan(PICTURE_HEIGHT);
  });

  test("a narrower square picture keeps side wrapping", () => {
    const { zones, wrapped } = laidOutCell(picture({ width: 100, distRight: 12 }));

    expect(zones).toEqual([{ leftMargin: 112, rightMargin: 0, topY: 0, bottomY: PICTURE_HEIGHT }]);
    expect(firstLineTop(wrapped)).toBe(0);
  });

  test.each<ImageRun["wrapType"]>(["behind", "inFront"])(
    "a wrapNone (%s) picture excludes no text",
    (wrapType) => {
      const { zones, wrapped, measure } = laidOutCell(picture({ wrapType }));

      expect(zones).toEqual([]);
      expect(firstLineTop(wrapped)).toBe(0);
      expect(measure.rows[0]!.height).toBeLessThan(PICTURE_HEIGHT);
    },
  );
});
