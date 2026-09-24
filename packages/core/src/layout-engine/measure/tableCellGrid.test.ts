import { describe, expect, test } from "bun:test";

import type { TableCell, TableRow } from "../types";
import {
  buildTableCellGrid,
  buildTableCellPlacements,
  getFirstAvailableColumn,
  getSourceCellAt,
  getSourceCellColumn,
  getTableCellSpacingInsetsY,
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

  test("lets a styled zero-width hairline own a shared edge without adding height", () => {
    const above = cell(1, { borders: { bottom: { width: 0, style: "solid" } } });
    const below = cell(2, {
      borders: {
        top: { width: 3, style: "solid" },
        bottom: { width: 4, style: "solid" },
      },
    });
    const grid = buildTableCellGrid([row(10, [above]), row(11, [below])], 1);

    expect(getTableCellVerticalBorderHeight(grid, below, 1)).toBe(4);
  });
});

describe("w:tblCellSpacing cell boxes", () => {
  test("insets each cell one unit inside its slot and one more at the table edge", () => {
    const first = cell(1);
    const middle = cell(2);
    const last = cell(3);
    const grid = buildTableCellGrid([row(10, [first, middle, last])], 3);
    const placements = buildTableCellPlacements({
      grid,
      columnWidths: [100, 80, 120],
      bidi: false,
      cellSpacing: 3,
    });

    expect(placements.get(first)).toMatchObject({ left: 6, width: 91 });
    expect(placements.get(middle)).toMatchObject({ left: 103, width: 74 });
    expect(placements.get(last)).toMatchObject({ left: 183, width: 111 });
  });

  test("mirrors the spaced boxes for a right-to-left table", () => {
    const first = cell(1);
    const last = cell(2);
    const grid = buildTableCellGrid([row(10, [first, last])], 2);
    const placements = buildTableCellPlacements({
      grid,
      columnWidths: [100, 60],
      bidi: true,
      cellSpacing: 2,
    });

    expect(placements.get(first)).toMatchObject({ left: 62, width: 94 });
    expect(placements.get(last)).toMatchObject({ left: 4, width: 54 });
  });

  test("keeps a row that stops short of the grid one unit from the next slot", () => {
    const leading = cell(1);
    const grid = buildTableCellGrid([{ id: 10, cells: [leading], gridAfter: 1 }], 2);
    const placements = buildTableCellPlacements({
      grid,
      columnWidths: [100, 100],
      bidi: false,
      cellSpacing: 2,
    });

    expect(placements.get(leading)).toMatchObject({ left: 4, width: 94 });
  });

  test("leaves collapsed tables on their grid lines", () => {
    const only = cell(1);
    const grid = buildTableCellGrid([row(10, [only])], 1);

    expect(
      buildTableCellPlacements({ grid, columnWidths: [90], bidi: false, cellSpacing: 0 }).get(only),
    ).toMatchObject({ left: 0, width: 90 });
    expect(getTableCellSpacingInsetsY(undefined, 0, 1, 1)).toEqual({ top: 0, bottom: 0 });
  });

  test("adds the table-edge unit above the first row and below the last", () => {
    expect(getTableCellSpacingInsetsY(2, 0, 1, 3)).toEqual({ top: 4, bottom: 2 });
    expect(getTableCellSpacingInsetsY(2, 1, 1, 3)).toEqual({ top: 2, bottom: 2 });
    expect(getTableCellSpacingInsetsY(2, 1, 2, 3)).toEqual({ top: 2, bottom: 4 });
  });

  test("counts both horizontal borders of a spaced cell", () => {
    const above = cell(1, { borders: { bottom: { width: 2, style: "solid" } } });
    const below = cell(2, {
      borders: {
        top: { width: 3, style: "solid" },
        bottom: { width: 4, style: "solid" },
      },
    });
    const grid = buildTableCellGrid([row(10, [above]), row(11, [below])], 1);

    expect(getTableCellVerticalBorderHeight(grid, below, 1)).toBe(4);
    expect(getTableCellVerticalBorderHeight(grid, below, 1, true)).toBe(7);
  });
});
