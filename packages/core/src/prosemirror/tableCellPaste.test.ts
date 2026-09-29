import { describe, expect, test } from "bun:test";
import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import { EditorState as PMEditorState, TextSelection, type EditorState } from "prosemirror-state";
import { CellSelection, TableMap } from "prosemirror-tables";

import { documentShape } from "../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  resolveAllChanges,
  saveHarnessState,
  summarizeState,
  textblocks,
} from "../__tests__/editorHarness";
import type { Document } from "../types/document";
import { schema as documentSchema } from "./schema";
import { pasteTableCells } from "./tableCellPaste";

/**
 * The shape's table follows a paragraph, so it does not start the document:
 * row 1 holds a 2×2 merge (A1:B2), "Cell C1" and a vertical merge D1:D2; row 2
 * holds only "Cell C2"; A3:A4 is merged too.
 */
const loadTables = async (): Promise<Document> =>
  parseShapeDocument(await documentShape("tables").build());

const outerTable = (doc: PMNode): { table: PMNode; pos: number } => {
  let found: { table: PMNode; pos: number } | null = null;
  doc.descendants((node, pos) => {
    if (found) {
      return false;
    }
    if (node.type.name === "table") {
      found = { table: node, pos };
      return false;
    }
    return true;
  });
  if (!found) {
    throw new Error("No table");
  }
  return found;
};

/** Each row's cells as `text@row,column rowspan×colspan`. */
const layout = (doc: PMNode): string[][] => {
  const { table } = outerTable(doc);
  const map = TableMap.get(table);
  expect(map.problems ?? []).toEqual([]);
  expect(table.attrs["columnWidths"]).toHaveLength(map.width);
  const rows: string[][] = [];
  let offset = 1;
  table.forEach((row, _rowOffset, rowIndex) => {
    const cells: string[] = [];
    let cellOffset = offset + 1;
    row.forEach((cell) => {
      const rect = map.findCell(cellOffset - 1);
      cells.push(
        `${cell.textContent}@${rowIndex},${rect.left} ${cell.attrs["rowspan"]}×${cell.attrs["colspan"]}`,
      );
      cellOffset += cell.nodeSize;
    });
    rows.push(cells);
    offset += row.nodeSize;
  });
  return rows;
};

const pastedTable = (state: EditorState): Slice => {
  const { schema } = state;
  const cell = (text: string) =>
    schema.node("tableCell", null, [schema.node("paragraph", null, schema.text(text))]);
  return new Slice(
    Fragment.from(
      schema.node("table", null, [
        schema.node("tableRow", null, [cell("P1"), cell("P2")]),
        schema.node("tableRow", null, [cell("P3"), cell("P4")]),
      ]),
    ),
    0,
    0,
  );
};

const pasteAtCell = (base: Document, mode: "editing" | "suggesting", text: string) => {
  const state = placeSelection(createHarnessState(base, mode), text, "caret-middle");
  if (!state) {
    throw new Error(`No caret in ${text}`);
  }
  const view = new HeadlessEditorView(state);
  view.paste(pastedTable(view.state));
  return view;
};

describe("pasting table cells into a merged table", () => {
  test("splits the merge a pasted block cuts through and keeps the grid tiled", async () => {
    const base = await loadTables();
    const view = pasteAtCell(base, "editing", "Cell C2");

    expect(layout(view.state.doc)).toEqual([
      ["Merged A1@0,0 2×2", "Cell C1@0,2 1×1", "Cell D1@0,3 1×1"],
      ["P1@1,2 1×1", "P2@1,3 1×1"],
      ["Cell A3@2,0 2×1", "Cell B3@2,1 1×1", "P3@2,2 1×1", "P4@2,3 1×1"],
      ["Cell B4@3,1 1×1", "Cell C4@3,2 1×1", "Cell D4@3,3 1×1"],
    ]);
    expect(view.state.selection).toBeInstanceOf(CellSelection);
    await saveHarnessState(view.state, base);
  });

  test("grows the table and its grid when the block reaches past the last column", async () => {
    const base = await loadTables();
    const view = pasteAtCell(base, "editing", "Cell D4");

    const rows = layout(view.state.doc);
    expect(rows).toHaveLength(5);
    expect(rows[3]).toContain("P1@3,3 1×1");
    expect(rows[3]).toContain("P2@3,4 1×1");
    expect(rows[4]).toContain("P4@4,4 1×1");
    await saveHarnessState(view.state, base);
  });

  test("tracks the paste in suggesting mode, so accepting matches editing and rejecting restores", async () => {
    const base = await loadTables();
    const editing = pasteAtCell(base, "editing", "Cell C2");
    const suggesting = pasteAtCell(base, "suggesting", "Cell C2");
    const original = createHarnessState(base, "suggesting");

    expect(summarizeState(resolveAllChanges(suggesting.state, "reject"))).toEqual(
      summarizeState(original),
    );
    expect(summarizeState(resolveAllChanges(suggesting.state, "accept"))).toEqual(
      summarizeState(editing.state),
    );
    expect(layout(resolveAllChanges(suggesting.state, "reject").doc)).toEqual(layout(original.doc));
    await saveHarnessState(suggesting.state, base);
  });

  test("tracks the rows and cells a tracked paste grows the table by", async () => {
    const base = await loadTables();
    const editing = pasteAtCell(base, "editing", "Cell D4");
    const suggesting = pasteAtCell(base, "suggesting", "Cell D4");
    const original = createHarnessState(base, "suggesting");

    const rejected = resolveAllChanges(suggesting.state, "reject");
    expect(layout(rejected.doc)).toEqual(layout(original.doc));
    expect(outerTable(rejected.doc).table.attrs["columnWidths"]).toEqual(
      outerTable(original.doc).table.attrs["columnWidths"],
    );
    expect(layout(resolveAllChanges(suggesting.state, "accept").doc)).toEqual(
      layout(editing.state.doc),
    );
    await saveHarnessState(suggesting.state, base);
  });

  test("fills every cell of a cell selection without changing the table's shape", async () => {
    const base = await loadTables();
    const state = placeSelection(createHarnessState(base, "editing"), "Cell C2", "cross-paragraph");
    if (!(state?.selection instanceof CellSelection)) {
      throw new Error("Expected a cell selection");
    }
    const before = layout(state.doc).map((row) => row.map((cell) => cell.replace(/^.*@/u, "@")));
    const view = new HeadlessEditorView(state);
    view.paste(new Slice(Fragment.from(view.state.schema.text("Filled")), 0, 0));

    const after = layout(view.state.doc);
    expect(after.map((row) => row.map((cell) => cell.replace(/^.*@/u, "@")))).toEqual(before);
    const filled = textblocks(view.state.doc).filter(({ node }) => node.textContent === "Filled");
    expect(filled.length).toBeGreaterThan(1);
    await saveHarnessState(view.state, base);
  });

  test.each([
    ["into a cell", "editing", "caret-end"],
    ["over everything, suggesting", "suggesting", "document"],
  ] as const)(
    "a pasted copy of a merged cell does not claim the stored continuation of its source (%s)",
    async (_label, mode, placement) => {
      const base = await loadTables();
      const state = placeSelection(createHarnessState(base, mode), "Cell C2", placement);
      if (!state) {
        throw new Error("No caret");
      }
      const view = new HeadlessEditorView(state);
      const blocks = textblocks(view.state.doc);
      const index = blocks.findIndex(({ node }) => node.textContent === "Cell C2");
      const first = blocks[index];
      const second = blocks[index + 1];
      if (!first || !second) {
        throw new Error("No blocks to copy");
      }
      // "Cell C2" through "Cell A3", whose cell stores the continuation of A3:A4.
      view.paste(view.state.doc.slice(first.pos + 1, second.pos + 1 + second.node.content.size));

      layout(view.state.doc);
      await saveHarnessState(view.state, base);
    },
  );

  test("splits a cell merged across the block's left edge", () => {
    const cell = (text: string, colspan = 1) =>
      documentSchema.node("tableCell", { colspan }, [
        documentSchema.node("paragraph", null, documentSchema.text(text)),
      ]);
    const doc = documentSchema.node("doc", null, [
      documentSchema.node("table", { columnWidths: [1000, 1000, 1000] }, [
        documentSchema.node("tableRow", null, [cell("x"), cell("y"), cell("z")]),
        documentSchema.node("tableRow", null, [cell("A", 2), cell("C")]),
      ]),
    ]);
    let caret = 0;
    doc.descendants((node, pos) => {
      if (node.isText && node.text === "y") {
        caret = pos;
      }
    });
    let state = PMEditorState.create({ doc, selection: TextSelection.create(doc, caret) });
    expect(
      pasteTableCells(state, pastedTable(state), (tr) => {
        state = state.apply(tr);
      }),
    ).toBe(true);

    expect(layout(state.doc)).toEqual([
      ["x@0,0 1×1", "P1@0,1 1×1", "P2@0,2 1×1"],
      ["A@1,0 1×1", "P3@1,1 1×1", "P4@1,2 1×1"],
    ]);
  });
});
