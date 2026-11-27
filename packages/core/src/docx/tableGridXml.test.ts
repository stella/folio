import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Table } from "../types/document";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { parseTableGeometryChange } from "./reviewResolutionProvenance";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
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

const readTable = (body: string) => {
  const xml = body.replace(
    "<w:tbl>",
    '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">',
  );
  const element = parseXmlDocument(xml);
  if (element === null) throw new TypeError("Expected a table element.");
  const table = parseTable(element, new Map(), null, null, null, null);
  if (table === undefined) throw new TypeError("Expected a parsed table.");
  return table;
};

test.each([
  ["cellDel", REVISION_DECISIONS.ACCEPT],
  ["cellIns", REVISION_DECISIONS.REJECT],
] as const)(
  "review of %s can leave inferred grid columns",
  (mark, decision) => {
    const table = readTable(
      [
        '<w:tbl><w:tblGrid><w:gridCol w:w="600"/></w:tblGrid><w:tr>',
        `<w:tc><w:tcPr><w:${mark} w:id="7" w:author="Editor"/></w:tcPr>`,
        '<w:p w14:paraId="00000001"><w:r><w:t>Reviewed</w:t></w:r></w:p></w:tc>',
        '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>',
        '<w:p w14:paraId="00000002"><w:r><w:t>Retained</w:t></w:r></w:p></w:tc>',
        "</w:tr></w:tbl>",
      ].join(""),
    );
    expect(table.columnWidths).toEqual([600]);
    const document = { package: { document: { content: [table] } } };
    const resolved = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [7],
      decision,
    }).unwrap();
    const result = resolved.document.package.document.content.at(0);
    if (result?.type !== "table")
      throw new TypeError("Expected a retained table.");
    expect(result.columnWidths).toBeUndefined();
    expect(gridColumnCount(result)).toBe(2);
    expect(result.rows.at(0)?.cells).toEqual(table.rows.at(0)?.cells.slice(1));
    const saved = serializeTable(result, serializeParagraph);
    expect(saved).toContain("<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>");
    expect(saved).toContain("Retained");
    expect(saved).not.toContain("Reviewed");
    const reopened = readTable(saved);
    expect(gridColumnCount(reopened)).toBe(2);
    expect(reopened.rows.at(0)?.cells).toHaveLength(1);
    expect(
      applyDocumentOps(resolved.document, resolved.inverse).unwrap().document,
    ).toStrictEqual(document);
  },
);

test.each(["ins", "del"] as const)(
  "native row review preserves a fully omitted row (%s)",
  (mark) => {
    const table = readTable(
      [
        '<w:tbl><w:tblGrid><w:gridCol w:w="600"/><w:gridCol w:w="700"/></w:tblGrid>',
        '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p w14:paraId="00000001"/></w:tc></w:tr>',
        `<w:tr><w:trPr><w:gridBefore w:val="2"/><w:${mark} w:id="7" w:author="Editor"/></w:trPr></w:tr>`,
        "</w:tbl>",
      ].join(""),
    );
    const document = { package: { document: { content: [table] } } };
    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const resolved = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [7],
        decision,
      }).unwrap();
      const result = resolved.document.package.document.content.at(0);
      if (result?.type !== "table")
        throw new TypeError("Expected a retained table.");
      const keepsRow =
        (mark === "ins") === (decision === REVISION_DECISIONS.ACCEPT);
      const reopened = readTable(serializeTable(result, serializeParagraph));
      expect(reopened.rows).toHaveLength(keepsRow ? 2 : 1);
      if (keepsRow) {
        expect(reopened.rows.at(1)?.cells).toEqual([]);
        expect(reopened.rows.at(1)?.formatting?.gridBefore).toBe(2);
        expect(reopened.rows.at(1)?.structuralChange).toBeUndefined();
      }
      expect(
        applyDocumentOps(resolved.document, resolved.inverse).unwrap().document,
      ).toStrictEqual(document);
    }
  },
);

test.each([
  { cut: 0, last: 1, expected: [undefined, undefined] },
  { cut: 0, last: 2, expected: [undefined, "restart", "continue"] },
  { cut: 1, last: 3, expected: ["restart", undefined, "restart", "continue"] },
])(
  "selective cell review retains vertical merge fragments ($cut of $last)",
  ({ cut, last, expected }) => {
    const table = readTable(
      [
        '<w:tbl><w:tblGrid><w:gridCol w:w="600"/><w:gridCol w:w="700"/></w:tblGrid>',
        ...Array.from({ length: last + 1 }, (_value, index) => {
          const merge =
            index === 0 ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>";
          const change =
            index === cut ? '<w:cellDel w:id="7" w:author="Editor"/>' : "";
          return `<w:tr><w:tc><w:tcPr>${merge}${change}</w:tcPr><w:p w14:paraId="0000001${index}"><w:r><w:t>Merge cell ${index}</w:t></w:r></w:p></w:tc><w:tc><w:p w14:paraId="0000002${index}"/></w:tc></w:tr>`;
        }),
        "</w:tbl>",
      ].join(""),
    );
    const document = { package: { document: { content: [table] } } };
    const resolved = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [7],
      decision: REVISION_DECISIONS.ACCEPT,
    }).unwrap();
    const result = resolved.document.package.document.content.at(0);
    if (result?.type !== "table")
      throw new TypeError("Expected a retained table.");
    expect(result.rows).toHaveLength(last + 1);
    const mergedCells = (source: Table) =>
      source.rows.map((row) => row.cells.at(0)?.formatting?.vMerge);
    expect(mergedCells(result)).toEqual(expected);
    const saved = serializeTable(result, serializeParagraph);
    expect(saved).not.toContain(`Merge cell ${cut}`);
    for (let index = 0; index <= last; index++)
      if (index !== cut) expect(saved).toContain(`Merge cell ${index}`);
    expect(mergedCells(readTable(saved))).toEqual(expected);
    expect(
      applyDocumentOps(resolved.document, resolved.inverse).unwrap().document,
    ).toStrictEqual(document);
  },
);
