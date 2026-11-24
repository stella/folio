import { expect, test } from "bun:test";

import type { Table } from "../types/document";
import {
  gridColumnCount,
  MAX_GRID_COLUMNS,
  modelTableGrid,
} from "./tableGridXml";

test("a source's skipped-column count cannot expand the grid without bound", () => {
  const table: Table = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        formatting: { gridBefore: 1_000_000_000 },
        cells: [{ type: "tableCell", content: [] }],
      },
    ],
  };
  expect(gridColumnCount(table)).toBe(MAX_GRID_COLUMNS);
  expect(modelTableGrid(table).match(/<w:gridCol\/>/gu)).toHaveLength(
    MAX_GRID_COLUMNS,
  );
});
