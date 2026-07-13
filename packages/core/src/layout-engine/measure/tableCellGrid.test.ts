import { describe, expect, test } from "bun:test";

import type { TableCell, TableRow } from "../types";
import {
  buildTableCellGrid,
  getFirstAvailableColumn,
  getSourceCellAt,
  getSourceCellColumn,
  getTableCellVerticalBorderHeight,
} from "./tableCellGrid";

const cell = (id: number, options: Partial<TableCell> = {}): TableCell => ({
  id,
  blocks: [],
  ...options,
});

const row = (id: number, cells: TableCell[], gridBefore?: number): TableRow => ({
  id,
  cells,
  ...(gridBefore === undefined ? {} : { gridBefore }),
});

describe("table cell grid", () => {
  test("indexes omitted columns and vertically merged cells once", () => {
    const merged = cell(1, { colSpan: 2, rowSpan: 2 });
    const firstRowTail = cell(2);
    const secondRowTail = cell(3);
    const grid = buildTableCellGrid(
      [row(10, [merged, firstRowTail], 1), row(11, [secondRowTail], 1)],
      4,
    );

    expect(getSourceCellColumn(grid, merged)).toBe(1);
    expect(getSourceCellColumn(grid, firstRowTail)).toBe(3);
    expect(getSourceCellColumn(grid, secondRowTail)).toBe(3);
    expect(getSourceCellAt(grid, 1, 1)).toBe(merged);
    expect(getSourceCellAt(grid, 1, 2)).toBe(merged);
    expect(getSourceCellAt(grid, 1, 3)).toBe(secondRowTail);
    expect(getFirstAvailableColumn(grid, 1, 1)).toBe(3);
  });

  test("does not index cell spans beyond the declared grid", () => {
    const spanning = cell(1, { colSpan: 3 });
    const grid = buildTableCellGrid([row(10, [spanning], 1)], 2);

    expect(getSourceCellAt(grid, 0, 1)).toBe(spanning);
    expect(getSourceCellAt(grid, 0, 2)).toBeUndefined();
  });

  test("assigns a shared horizontal edge to the cell above", () => {
    const above = cell(1, { borders: { bottom: { width: 2, style: "solid" } } });
    const below = cell(2, {
      borders: {
        top: { width: 3, style: "solid" },
        bottom: { width: 4, style: "solid" },
      },
    });
    const grid = buildTableCellGrid([row(10, [above]), row(11, [below])], 1);

    expect(getTableCellVerticalBorderHeight(grid, below, 1)).toBe(4);
  });

  test("retains the lower cell edge when the border above is paintless", () => {
    const above = cell(1, { borders: { bottom: { width: 2, style: "none" } } });
    const below = cell(2, {
      borders: {
        top: { width: 3, style: "solid" },
        bottom: { width: 4, style: "solid" },
      },
    });
    const grid = buildTableCellGrid([row(10, [above]), row(11, [below])], 1);

    expect(getTableCellVerticalBorderHeight(grid, below, 1)).toBe(7);
  });
});
