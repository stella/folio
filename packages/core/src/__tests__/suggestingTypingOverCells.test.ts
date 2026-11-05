/**
 * Text typed over selected cells in suggesting mode: a cell's selected range
 * spans the block boundaries inside the cell, but the text lands in its paragraph.
 */

import { describe, expect, test } from "bun:test";
import { NodeSelection } from "prosemirror-state";
import type { EditorState } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { shapeArrayBuffer } from "./documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
} from "./editorHarness";

const tablesState = async () =>
  createHarnessState(
    await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("tables"))),
    "suggesting",
  );

const selectFirstCells = (state: EditorState) => {
  const cells: number[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name === "tableCell") cells.push(pos);
    return true;
  });
  const [anchor, head] = cells;
  if (anchor === undefined || head === undefined)
    throw new Error("Fixture has no adjacent cells");
  return state.apply(
    state.tr.setSelection(CellSelection.create(state.doc, anchor, head)),
  );
};

const cellParagraphs = (state: EditorState, cellPos: number) => {
  const cell = state.doc.nodeAt(cellPos);
  if (!cell) throw new Error("Cell is gone");
  const texts: string[] = [];
  cell.forEach((child) => {
    texts.push(child.textContent);
  });
  return texts;
};

describe("typing over selected cells while suggesting", () => {
  for (const deletesFirst of [true, false]) {
    test(`keeps the typed text in order, in the cell's paragraph (${deletesFirst ? "after Delete" : "directly"})`, async () => {
      const state = selectFirstCells(await tablesState());
      if (!(state.selection instanceof CellSelection))
        throw new Error("Fixture did not select cells");
      const cellPos = state.selection.$headCell.pos;
      const [original] = cellParagraphs(state, cellPos);
      const view = new HeadlessEditorView(state);
      if (deletesFirst) view.pressKey("Delete");
      view.typeText("alpha");

      expect(cellParagraphs(view.state, cellPos)).toEqual([`${original}alpha`]);
      const { $head } = view.state.selection;
      expect(view.state.selection.empty).toBe(true);
      expect($head.parent.textContent.slice(0, $head.parentOffset)).toBe(
        `${original}alpha`,
      );
      expect(
        cellParagraphs(resolveAllChanges(view.state, "accept"), cellPos),
      ).toEqual(["alpha"]);
      expect(
        cellParagraphs(resolveAllChanges(view.state, "reject"), cellPos),
      ).toEqual([original]);
    });
  }
});

describe("typing over a selected block while suggesting", () => {
  test("keeps the typed text in order, tracked, with the caret after it", async () => {
    const state = await tablesState();
    let tablePos: number | null = null;
    state.doc.forEach((node, offset) => {
      if (node.type.name === "table") tablePos = offset;
    });
    if (tablePos === null) throw new Error("Fixture has no table");
    const view = new HeadlessEditorView(
      state.apply(
        state.tr.setSelection(NodeSelection.create(state.doc, tablePos)),
      ),
    );
    view.typeText("alpha");

    const { $head } = view.state.selection;
    expect($head.parent.textContent).toBe("alpha");
    expect($head.parentOffset).toBe(5);
    expect(
      resolveAllChanges(view.state, "reject").doc.textContent,
    ).not.toContain("alpha");
  });
});
