import { expect, test } from "bun:test";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Table } from "../types/document";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { parseTableGeometryChange } from "./reviewResolutionProvenance";
import {
  gridColumnCount,
  MAX_GRID_COLUMNS,
  modelTableGrid,
} from "./tableGridXml";
import { parseTable } from "./tableParser";
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

test("a rebuilt table distinguishes a missing width from an explicit zero", () => {
  const widths = '<w:gridCol w:w="600"/><w:gridCol/><w:gridCol w:w="0"/>';
  const cells = ["00000001", "00000002", "00000003"]
    .map((id) => `<w:tc><w:p w14:paraId="${id}"/></w:tc>`)
    .join("");
  const input = parseXmlDocument(
    '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
      `<w:tblGrid>${widths}</w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`,
  );
  if (input === null) throw new TypeError("Expected a table XML element.");
  const table = parseTable(input, new Map(), null, null, null, null);
  if (table === undefined) throw new TypeError("Expected a table.");
  expect(table.omittedColumnWidthIndices).toEqual([1]);
  const document = {
    package: { document: { content: [table] } },
  } satisfies Document;
  const rebuilt = fromProseDoc(toProseDoc(document), document, {
    stylesheetSource: { type: "package" },
    reuse: "none",
  }).package.document.content.at(0);
  if (rebuilt?.type !== "table")
    throw new TypeError("Expected a rebuilt table.");
  expect(rebuilt.omittedColumnWidthIndices).toEqual([1]);
  expect(modelTableGrid(rebuilt)).toContain(widths);
});
