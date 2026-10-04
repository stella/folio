import { describe, expect, test } from "bun:test";

import type { Document, Table } from "../packages/docx-core/src/model/document";
import { normalizeDocumentPackage } from "./lib/corpus-invariants/model-equality";
import { canonicalReviewBlocks } from "../test/reviewProjection";

const info = (id: number) => ({ id, author: "Reviewer", date: "2026-05-06T07:08:09.000Z" });
const canonical = (table: Table) =>
  normalizeDocumentPackage({
    package: { document: { content: canonicalReviewBlocks([table]) } },
  } satisfies Document);
const tableWithHistory = (): Table => ({
  type: "table",
  columnWidths: [100],
  formatting: { justification: "center", gridChange: { id: 10, columnWidths: [90] } },
  propertyChanges: [{ type: "tablePropertyChange", info: info(1) }],
  rows: [
    {
      type: "tableRow",
      propertyChanges: [{ type: "tableRowPropertyChange", info: info(2) }],
      tablePropertyExceptions: { width: { type: "dxa", value: 100 } },
      tablePropertyExceptionChanges: [{ type: "tablePropertyExceptionChange", info: info(3) }],
      cells: [
        {
          type: "tableCell",
          content: [{ type: "paragraph", paraId: "00000001", content: [] }],
          propertyChanges: [{ type: "tableCellPropertyChange", info: info(4) }],
        },
      ],
    },
  ],
});

describe("canonical review projection for table property histories", () => {
  test("current captures and XML-only snapshots match their omitted equivalents", () => {
    const omitted = tableWithHistory();
    const captured = tableWithHistory();
    const capturedRow = captured.rows.at(0);
    const omittedRow = omitted.rows.at(0);
    const capturedCell = capturedRow?.cells.at(0);
    const omittedCell = omittedRow?.cells.at(0);
    if (
      capturedRow === undefined ||
      omittedRow === undefined ||
      capturedCell === undefined ||
      omittedCell === undefined
    )
      throw new Error("The property-history fixture is complete.");

    captured.formatting = {
      ...captured.formatting,
      sourceXml: "<w:tblPr/>",
      gridSourceXml: "<w:tblGrid><w:gridCol/></w:tblGrid>",
    };
    captured.propertyChanges = [
      {
        type: "tablePropertyChange",
        info: info(1),
        previousFormatting: {
          sourceXml: "<w:tblPr/>",
          gridSourceXml: "<w:tblGrid/>",
        },
        currentFormatting: captured.formatting,
      },
    ];
    capturedRow.formatting = { sourceXml: "<w:trPr/>" };
    capturedRow.propertyChanges = [
      {
        type: "tableRowPropertyChange",
        info: info(2),
        previousFormatting: {},
        currentFormatting: capturedRow.formatting,
      },
    ];
    capturedRow.tablePropertyExceptions = {
      width: { type: "dxa", value: 100 },
      sourceXml: "<w:tblPrEx/>",
    };
    capturedRow.tablePropertyExceptionChanges = [
      {
        type: "tablePropertyExceptionChange",
        info: info(3),
        previousFormatting: { sourceXml: "<w:tblPrEx/>" },
        currentFormatting: capturedRow.tablePropertyExceptions,
      },
    ];
    capturedCell.formatting = { sourceXml: "<w:tcPr/>" };
    capturedCell.propertyChanges = [
      {
        type: "tableCellPropertyChange",
        info: info(4),
        previousFormatting: { sourceXml: "<w:tcPr/>" },
        currentFormatting: capturedCell.formatting,
      },
    ];

    expect(canonical(captured)).toEqual(canonical(omitted));
  });

  test("authored table, row, cell, exception, grid, and history changes remain visible", () => {
    const baseline = tableWithHistory();
    baseline.propertyChanges = [
      {
        type: "tablePropertyChange",
        info: info(1),
        previousFormatting: { justification: "start" },
      },
    ];
    const row = baseline.rows.at(0);
    const cell = row?.cells.at(0);
    if (row === undefined || cell === undefined)
      throw new Error("The fixture has one row and cell.");
    row.formatting = { cantSplit: true };
    row.propertyChanges = [
      {
        type: "tableRowPropertyChange",
        info: info(2),
        previousFormatting: { gridBefore: 1 },
      },
    ];
    row.tablePropertyExceptionChanges = [
      {
        type: "tablePropertyExceptionChange",
        info: info(3),
        previousFormatting: { width: { type: "dxa", value: 90 } },
      },
    ];
    cell.formatting = { verticalAlign: "top" };
    cell.propertyChanges = [
      {
        type: "tableCellPropertyChange",
        info: info(4),
        previousFormatting: { verticalAlign: "bottom" },
      },
    ];

    const gridChanged = structuredClone(baseline);
    gridChanged.formatting = {
      ...gridChanged.formatting,
      gridChange: { id: 10, columnWidths: [80] },
    };
    expect(canonical(gridChanged)).not.toEqual(canonical(baseline));

    const previousGridHistoryChanged = structuredClone(baseline);
    const previousGridTableChange = previousGridHistoryChanged.propertyChanges?.at(0);
    if (previousGridTableChange === undefined) throw new Error("The table history exists.");
    previousGridTableChange.previousFormatting = {
      gridChange: { id: 11, columnWidths: [80] },
    };
    expect(canonical(previousGridHistoryChanged)).not.toEqual(canonical(baseline));

    const tableFormattingChanged = structuredClone(baseline);
    tableFormattingChanged.formatting = {
      ...tableFormattingChanged.formatting,
      justification: "start",
    };
    expect(canonical(tableFormattingChanged)).not.toEqual(canonical(baseline));

    const rowFormattingChanged = structuredClone(baseline);
    const changedRow = rowFormattingChanged.rows.at(0);
    if (changedRow === undefined) throw new Error("The row exists.");
    changedRow.formatting = { cantSplit: false };
    expect(canonical(rowFormattingChanged)).not.toEqual(canonical(baseline));

    const exceptionFormattingChanged = structuredClone(baseline);
    const changedExceptions = exceptionFormattingChanged.rows.at(0);
    if (changedExceptions === undefined) throw new Error("The row exists.");
    changedExceptions.tablePropertyExceptions = { width: { type: "dxa", value: 90 } };
    expect(canonical(exceptionFormattingChanged)).not.toEqual(canonical(baseline));

    const tableHistoryChanged = structuredClone(baseline);
    const tableChange = tableHistoryChanged.propertyChanges?.at(0);
    if (tableChange === undefined) throw new Error("The table history exists.");
    tableChange.previousFormatting = { justification: "end" };
    expect(canonical(tableHistoryChanged)).not.toEqual(canonical(baseline));

    const currentCaptureChanged = structuredClone(baseline);
    const currentChange = currentCaptureChanged.propertyChanges?.at(0);
    if (currentChange === undefined) throw new Error("The table history exists.");
    currentChange.currentFormatting = { justification: "end" };
    expect(canonical(currentCaptureChanged)).not.toEqual(canonical(baseline));

    const rowHistoryChanged = structuredClone(baseline);
    const rowChange = rowHistoryChanged.rows.at(0)?.propertyChanges?.at(0);
    if (rowChange === undefined) throw new Error("The row history exists.");
    rowChange.info.id += 10;
    expect(canonical(rowHistoryChanged)).not.toEqual(canonical(baseline));

    const exceptionHistoryChanged = structuredClone(baseline);
    const exceptionChange = exceptionHistoryChanged.rows
      .at(0)
      ?.tablePropertyExceptionChanges?.at(0);
    if (exceptionChange === undefined) throw new Error("The exception history exists.");
    exceptionChange.previousFormatting = { width: { type: "dxa", value: 80 } };
    expect(canonical(exceptionHistoryChanged)).not.toEqual(canonical(baseline));

    const cellFormattingChanged = structuredClone(baseline);
    const changedCell = cellFormattingChanged.rows.at(0)?.cells.at(0);
    if (changedCell === undefined) throw new Error("The cell exists.");
    changedCell.formatting = { verticalAlign: "center" };
    expect(canonical(cellFormattingChanged)).not.toEqual(canonical(baseline));
  });
});
