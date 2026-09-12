import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";

import {
  decodeTableCellParagraphSourcePayload,
  transportTableCellsWithParagraphPropertySources,
} from "../docx/paragraphPropertySource";
import { tableFromTemplate, tableRowFromTemplate } from "./table-template";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "text*", group: "block" },
    text: {},
    table: {
      content: "tableRow+",
      group: "block",
      tableRole: "table",
    },
    tableRow: {
      content: "tableCell*",
      tableRole: "row",
    },
    tableCell: {
      content: "block+",
      tableRole: "cell",
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        _preserveVMergeRestart: { default: null },
        _docxVMergeContinuationCells: { default: null },
      },
    },
  },
});

const continuationCells = () =>
  transportTableCellsWithParagraphPropertySources([
    { type: "tableCell", content: [{ type: "paragraph", content: [] }] },
  ]);

const mergedCell = () =>
  schema.node(
    "tableCell",
    {
      rowspan: 2,
      _docxVMergeContinuationCells: continuationCells(),
    },
    [schema.node("paragraph")],
  );

describe("table templates", () => {
  test("a row copy clears the continuation payload when it clamps the rowspan", () => {
    const copied = tableRowFromTemplate({
      template: schema.node("tableRow", null, [mergedCell()]),
      columnCount: 1,
    });

    expect(copied?.child(0).attrs["rowspan"]).toBe(1);
    expect(copied?.child(0).attrs["_docxVMergeContinuationCells"]).toBeNull();
  });

  test("a whole-table copy preserves the rowspan and its continuation payload", () => {
    const copied = tableFromTemplate({
      schema,
      template: schema.node("table", null, [
        schema.node("tableRow", null, [mergedCell()]),
        schema.node("tableRow"),
      ]),
    });
    const copiedCell = copied?.child(0).child(0);
    const payload = copiedCell?.attrs["_docxVMergeContinuationCells"];

    expect(copiedCell?.attrs["rowspan"]).toBe(2);
    expect(payload).not.toBeNull();
    expect(decodeTableCellParagraphSourcePayload(payload).cells).toHaveLength(1);
  });
});
