import { expect, test } from "bun:test";

import type { Table } from "../types/document";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { parseTableGeometryChange } from "./reviewResolutionProvenance";
import {
  gridColumnCount,
  MAX_GRID_COLUMNS,
  modelTableGrid,
} from "./tableGridXml";
import { parseXmlDocument } from "./xmlParser";

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

test.each(["deleteColumn", "mergeCells"] as const)(
  "the reader recognizes only supported %s geometry fields",
  (type) => {
    const geometry = { type, revisionIds: [50, 51] };
    const read = (value: object) => {
      const element = parseXmlDocument(
        `<w:tblGridChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}" folio:tableGeometryChange='${JSON.stringify({ version: 2, value })}'/>`,
      );
      if (element === null)
        throw new TypeError("Expected a grid change element.");
      return parseTableGeometryChange(element);
    };
    expect(read(geometry)).toEqual({ geometryChange: geometry });
    expect(read({ ...geometry, startColumn: 2 })).toEqual({});
  },
);
