import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { mergeTableRectangle, splitTableRectangle } from "./table-cell-mutations";

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
      content: "tableCell+",
      tableRole: "row",
    },
    tableCell: {
      content: "block+",
      tableRole: "cell",
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null },
        _preserveVMergeRestart: { default: null },
        _docxVMergeContinuationCells: { default: null },
      },
    },
  },
});

const cell = (text: string) =>
  schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text(text)])]);

const table = schema.node("table", null, [
  schema.node("tableRow", null, [cell("A"), cell("B"), cell("X")]),
  schema.node("tableRow", null, [cell("C"), cell("D"), cell("Y")]),
]);

describe("table cell mutations", () => {
  test("merges and splits one rectangle without losing cell content", () => {
    let state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const rectangle = { left: 0, top: 0, right: 2, bottom: 2 };
    const merged = mergeTableRectangle({
      tr: state.tr,
      tablePosition: 0,
      table: state.doc.child(0),
      rectangle,
    });
    if (!merged) {
      throw new Error("Expected the rectangle to merge.");
    }
    state = state.apply(merged);

    const mergedTable = state.doc.child(0);
    expect(TableMap.get(mergedTable)).toMatchObject({ width: 3, height: 2 });
    expect(mergedTable.firstChild?.firstChild?.attrs).toMatchObject({ colspan: 2, rowspan: 2 });
    expect(mergedTable.textContent).toBe("ABCDXY");

    const split = splitTableRectangle({
      tr: state.tr,
      tablePosition: 0,
      table: mergedTable,
      rectangle,
    });
    if (!split) {
      throw new Error("Expected the rectangle to split.");
    }
    state = state.apply(split);

    const splitTable = state.doc.child(0);
    expect(splitTable.childCount).toBe(2);
    expect(splitTable.child(0).childCount).toBe(3);
    expect(splitTable.child(1).childCount).toBe(3);
    expect(splitTable.textContent).toBe("ABCDXY");
  });

  test("rejects an invalid rectangle without adding transaction steps", () => {
    const state = EditorState.create({ schema, doc: schema.node("doc", null, [table]) });
    const tr = state.tr;

    expect(
      mergeTableRectangle({
        tr,
        tablePosition: 0,
        table: state.doc.child(0),
        rectangle: { left: -1, top: 0, right: 2, bottom: 2 },
      }),
    ).toBeNull();
    expect(tr.steps).toEqual([]);
  });
});
