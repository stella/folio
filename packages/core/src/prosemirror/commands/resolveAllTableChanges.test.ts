import { expect, test } from "bun:test";
import { Schema } from "prosemirror-model";

import { resolveAllTableChanges } from "./resolveAllTableChanges";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "text*", group: "block" },
    text: { group: "inline" },
    table: {
      content: "tableRow+",
      group: "block",
      tableRole: "table",
      attrs: {
        columnWidths: { default: null },
        _originalFormatting: { default: null },
      },
    },
    tableRow: {
      content: "tableCell*",
      tableRole: "row",
      attrs: { trIns: { default: null }, trDel: { default: null } },
    },
    tableCell: {
      content: "block+",
      tableRole: "cell",
      isolating: true,
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null },
        cellMarker: { default: null },
        _originalFormatting: { default: null },
        _docxVMergeContinuationCells: { default: null },
      },
    },
  },
});

const paragraph = (value: string) =>
  schema.node("paragraph", null, value ? [schema.text(value)] : []);
const cell = (value: string, attrs?: Record<string, unknown>) =>
  schema.node("tableCell", attrs, [paragraph(value)]);
const row = (cells: ReturnType<typeof cell>[], attrs?: Record<string, unknown>) =>
  schema.node("tableRow", attrs, cells);

test("batch row removal promotes a crossing cell and reduces its rowspan", () => {
  const table = schema.node("table", null, [
    row([cell("span", { rowspan: 2 }), cell("discard")], {
      trDel: { revisionId: 10, author: "Reviewer" },
    }),
    row([cell("keep")]),
  ]);
  const result = resolveAllTableChanges({ table, mode: "accept" });

  expect(result.failed).toBe(false);
  expect(result.node?.childCount).toBe(1);
  expect(result.node?.firstChild?.childCount).toBe(2);
  expect(result.node?.firstChild?.firstChild?.textContent).toBe("span");
  expect(result.node?.firstChild?.firstChild?.attrs["rowspan"]).toBe(1);
  expect(result.node?.firstChild?.child(1).textContent).toBe("keep");
  expect(result.paragraphOffsets).toHaveLength(2);
  expect(result.paragraphOffsets[1].source).toBeGreaterThan(result.paragraphOffsets[1].final);
  expect(result.positionMap.map(result.paragraphOffsets[1].source, 1)).toBe(
    result.paragraphOffsets[1].final,
  );
  expect(result.changedParagraphRanges).toEqual([
    {
      from: 2,
      to: 2 + (result.node?.firstChild?.firstChild?.nodeSize ?? 0),
    },
  ]);
});

test("batch cell removal preserves the surviving grid provenance", () => {
  const table = schema.node(
    "table",
    {
      columnWidths: [1200, 1800],
      _originalFormatting: { sourceXml: "<w:tblPr/>", gridSourceXml: "<w:tblGrid/>" },
    },
    [
      row([
        cell("left"),
        cell("deleted", {
          cellMarker: { kind: "del", info: { revisionId: 11 } },
        }),
      ]),
      row([
        cell("left too"),
        cell("deleted too", {
          cellMarker: { kind: "del", info: { revisionId: 12 } },
        }),
      ]),
    ],
  );
  const result = resolveAllTableChanges({ table, mode: "accept" });

  expect(result.failed).toBe(false);
  expect(result.node?.child(0).childCount).toBe(1);
  expect(result.node?.child(1).childCount).toBe(1);
  expect(result.node?.attrs["columnWidths"]).toEqual([1200]);
  expect(result.node?.attrs["_originalFormatting"]).toEqual({ sourceXml: "<w:tblPr/>" });
  expect(result.changedParagraphRanges).toEqual([{ from: 0, to: result.node?.nodeSize }]);
});

test("grid widths follow the cell whose removal actually shrinks the table", () => {
  const deleted = (value: string, id: number) =>
    cell(value, { cellMarker: { kind: "del", info: { revisionId: id } } });
  const table = schema.node("table", { columnWidths: [100, 200, 300] }, [
    row([cell("a"), cell("b"), deleted("c", 21)]),
    row([cell("d"), deleted("e", 22), cell("f")]),
  ]);
  const result = resolveAllTableChanges({ table, mode: "accept" });

  expect(result.failed).toBe(false);
  expect(result.node?.attrs["columnWidths"]).toEqual([100, 200]);
  const survivingParagraph = result.paragraphOffsets.find(({ source }) =>
    table.nodeAt(source)?.textContent.includes("f"),
  );
  expect(survivingParagraph).toBeDefined();
  if (survivingParagraph) {
    expect(result.positionMap.map(survivingParagraph.source, 1)).toBe(survivingParagraph.final);
  }
});

test("rejecting a visible merge restores one spanning cell", () => {
  const table = schema.node("table", null, [
    row([cell("top")]),
    row([
      cell("continuation", {
        cellMarker: {
          kind: "merge",
          info: { revisionId: 13, author: "Reviewer", date: null },
          verticalMergeOriginal: "continue",
        },
      }),
    ]),
  ]);
  const result = resolveAllTableChanges({ table, mode: "reject" });

  expect(result.failed).toBe(false);
  expect(result.node?.child(0).firstChild?.attrs["rowspan"]).toBe(2);
  expect(result.node?.child(1).childCount).toBe(0);
  expect(result.node?.child(0).firstChild?.attrs["_docxVMergeContinuationCells"]).toHaveLength(1);
});

test("rejecting a collapsed merge restores continuation cells", () => {
  const continuation = {
    type: "tableCell" as const,
    formatting: { vMerge: "continue" as const },
    structuralChange: {
      type: "tableCellMerge" as const,
      info: { id: 14, author: "Reviewer" },
      verticalMerge: "continue" as const,
    },
    content: [
      {
        _docxParagraphSourceBinding: { type: "authored" as const },
        type: "paragraph" as const,
        content: [],
      },
    ],
  };
  const table = schema.node("table", null, [
    row([cell("top", { rowspan: 2, _docxVMergeContinuationCells: [continuation] })]),
    row([]),
  ]);
  const result = resolveAllTableChanges({ table, mode: "reject" });

  expect(result.failed).toBe(false);
  expect(result.node?.child(0).firstChild?.attrs["rowspan"]).toBe(1);
  expect(result.node?.child(1).childCount).toBe(1);
  expect(result.node?.child(1).firstChild?.attrs["cellMarker"]).toBeNull();
});

test("row and merge revisions resolve in the same pass", () => {
  const table = schema.node("table", null, [
    row([cell("top")]),
    row([
      cell("continuation", {
        cellMarker: {
          kind: "merge",
          info: { revisionId: 15, author: "Reviewer", date: null },
          verticalMergeOriginal: "continue",
        },
      }),
    ]),
    row([cell("removed")], { trDel: { revisionId: 16, author: "Reviewer" } }),
  ]);
  const result = resolveAllTableChanges({ table, mode: "accept" });

  expect(result.failed).toBe(false);
  expect(result.node?.childCount).toBe(2);
  expect(result.node?.child(1).firstChild?.attrs["cellMarker"]).toBeNull();
});

test("a chain of visible splits preserves continuation order", () => {
  const rows = [row([cell("top")])];
  for (let index = 1; index <= 64; index++) {
    rows.push(
      row([
        cell(`continuation ${index}`, {
          cellMarker: {
            kind: "merge",
            info: { revisionId: 100 + index, author: "Reviewer", date: null },
            verticalMergeOriginal: "continue",
          },
        }),
      ]),
    );
  }
  const table = schema.node("table", null, rows);
  const result = resolveAllTableChanges({ table, mode: "reject" });

  expect(result.failed).toBe(false);
  expect(result.node?.firstChild?.firstChild?.attrs["rowspan"]).toBe(65);
  const continuationCells =
    result.node?.firstChild?.firstChild?.attrs["_docxVMergeContinuationCells"];
  expect(continuationCells).toHaveLength(64);
  expect(continuationCells[0].content[0].content[0].content[0].text).toBe("continuation 1");
  expect(continuationCells[63].content[0].content[0].content[0].text).toBe("continuation 64");
});
