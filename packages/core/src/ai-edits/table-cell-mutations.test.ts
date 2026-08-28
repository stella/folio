import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import { history, undo } from "prosemirror-history";
import { EditorState } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { standaloneTableCellFromProseMirror } from "../prosemirror/conversion/fromProseDoc";
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
        width: { default: null },
        widthType: { default: null },
        _originalFormatting: { default: null },
        _preserveVMergeRestart: { default: null },
        _docxVMergeContinuationCells: { default: null },
      },
    },
  },
});

const cell = (text: string) =>
  schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text(text)])]);

type WidthCellOptions = {
  text: string;
  width: number;
  widthType: "dxa" | "pct";
  colwidth?: number;
  preserveOriginalFormatting?: boolean;
};

const widthCell = ({
  text,
  width,
  widthType,
  colwidth,
  preserveOriginalFormatting = true,
}: WidthCellOptions) =>
  schema.node(
    "tableCell",
    {
      colwidth: colwidth === undefined ? null : [colwidth],
      width,
      widthType,
      _originalFormatting: preserveOriginalFormatting
        ? { width: { value: width, type: widthType } }
        : null,
    },
    [schema.node("paragraph", null, [schema.text(text)])],
  );

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

  test("sums compatible preferred widths and restores them on undo", () => {
    const bilingualRow = schema.node("table", null, [
      schema.node("tableRow", null, [
        widthCell({ text: "Source", width: 2500, widthType: "pct", colwidth: 4860 }),
        widthCell({ text: "Translation", width: 2500, widthType: "pct", colwidth: 4860 }),
      ]),
    ]);
    let state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [bilingualRow]),
      plugins: [history()],
    });
    const merged = mergeTableRectangle({
      tr: state.tr,
      tablePosition: 0,
      table: state.doc.child(0),
      rectangle: { left: 0, top: 0, right: 2, bottom: 1 },
    });
    if (!merged) {
      throw new Error("Expected the bilingual row to merge.");
    }
    state = state.apply(merged);

    const mergedCell = state.doc.child(0).child(0).child(0);
    expect(mergedCell.attrs).toMatchObject({
      colspan: 2,
      colwidth: [4860, 4860],
      width: 5000,
      widthType: "pct",
    });
    expect(standaloneTableCellFromProseMirror(mergedCell).formatting?.width).toEqual({
      value: 5000,
      type: "pct",
    });

    expect(
      undo(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);
    const restoredRow = state.doc.child(0).child(0);
    expect(restoredRow.childCount).toBe(2);
    expect(restoredRow.child(0).attrs).toMatchObject({
      colwidth: [4860],
      width: 2500,
      widthType: "pct",
    });
    expect(restoredRow.child(1).attrs).toMatchObject({
      colwidth: [4860],
      width: 2500,
      widthType: "pct",
    });
  });

  test("sums fixed preferred widths", () => {
    const fixedWidthRow = schema.node("table", null, [
      schema.node("tableRow", null, [
        widthCell({ text: "First", width: 2400, widthType: "dxa" }),
        widthCell({ text: "Second", width: 3600, widthType: "dxa" }),
      ]),
    ]);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [fixedWidthRow]),
    });
    const merged = mergeTableRectangle({
      tr: state.tr,
      tablePosition: 0,
      table: state.doc.child(0),
      rectangle: { left: 0, top: 0, right: 2, bottom: 1 },
    });
    if (!merged) {
      throw new Error("Expected the fixed-width row to merge.");
    }

    expect(merged.doc.child(0).child(0).child(0).attrs).toMatchObject({
      width: 6000,
      widthType: "dxa",
    });
  });

  test("clears incompatible preferred widths instead of retaining the first cell width", () => {
    const mixedWidthRow = schema.node("table", null, [
      schema.node("tableRow", null, [
        widthCell({ text: "First", width: 2500, widthType: "pct" }),
        widthCell({ text: "Second", width: 3600, widthType: "dxa" }),
      ]),
    ]);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [mixedWidthRow]),
    });
    const merged = mergeTableRectangle({
      tr: state.tr,
      tablePosition: 0,
      table: state.doc.child(0),
      rectangle: { left: 0, top: 0, right: 2, bottom: 1 },
    });
    if (!merged) {
      throw new Error("Expected the mixed-width row to merge.");
    }

    const mergedCell = merged.doc.child(0).child(0).child(0);
    expect(mergedCell.attrs).toMatchObject({ width: null, widthType: null });
    expect(standaloneTableCellFromProseMirror(mergedCell).formatting?.width).toBeUndefined();
  });

  test("omits an incompatible merged width when no original cell formatting exists", () => {
    const mixedWidthRow = schema.node("table", null, [
      schema.node("tableRow", null, [
        widthCell({
          text: "First",
          width: 2500,
          widthType: "pct",
          preserveOriginalFormatting: false,
        }),
        widthCell({
          text: "Second",
          width: 3600,
          widthType: "dxa",
          preserveOriginalFormatting: false,
        }),
      ]),
    ]);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [mixedWidthRow]),
    });
    const merged = mergeTableRectangle({
      tr: state.tr,
      tablePosition: 0,
      table: state.doc.child(0),
      rectangle: { left: 0, top: 0, right: 2, bottom: 1 },
    });
    if (!merged) {
      throw new Error("Expected the mixed-width row to merge.");
    }

    const mergedCell = merged.doc.child(0).child(0).child(0);
    expect(standaloneTableCellFromProseMirror(mergedCell).formatting).toEqual({ gridSpan: 2 });
  });
});
