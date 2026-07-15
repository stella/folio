import { describe, expect, test } from "bun:test";

import { TableSelectionManager } from "./TableSelectionManager";

describe("TableSelectionManager", () => {
  test("tracks coordinates without inventing document context", () => {
    const manager = new TableSelectionManager();

    manager.selectCellCoordinates({ tableIndex: 2, rowIndex: 3, columnIndex: 4 });

    expect(manager.getSnapshot()).toEqual({
      context: null,
      table: null,
      tableIndex: 2,
      rowIndex: 3,
      columnIndex: 4,
    });
    expect(manager.isCellSelected(2, 3, 4)).toBe(true);
  });
});
