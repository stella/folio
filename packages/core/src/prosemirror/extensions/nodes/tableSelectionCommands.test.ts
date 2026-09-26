import { describe, expect, test } from "bun:test";
import { CellSelection } from "prosemirror-tables";

import { documentShape } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  harnessManager,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
} from "../../../__tests__/editorHarness";

// The shape's outer table: three rows over a three-column grid, with a
// horizontal merge in row 1, a vertical merge down column 1 of rows 2–3 (one
// cell with a row span) and a nested table in row 2.
const caretInTable = async () => {
  const document = await parseShapeDocument(await documentShape("tables").build());
  const state = placeSelection(createHarnessState(document, "editing"), "Cell B1", "caret-middle");
  if (!state) {
    throw new Error("the shape has no focus paragraph");
  }
  return new HeadlessEditorView(state);
};

describe("table selection commands", () => {
  test("selectTable selects every cell of the table holding the caret", async () => {
    const view = await caretInTable();

    expect(harnessManager().requireCommand("selectTable")()(view.state, view.dispatch)).toBe(true);

    const { selection } = view.state;
    expect(selection).toBeInstanceOf(CellSelection);
    const cells: string[] = [];
    (selection as CellSelection).forEachCell((cell) => {
      cells.push(cell.firstChild?.textContent ?? "");
    });
    expect(cells).toEqual([
      "Merged A1",
      "Cell C1",
      "Cell A2",
      "Cell B1",
      "Outer",
      "Cell B3",
      "Cell C3",
    ]);
  });

  test("selectRow selects the caret's row, with the rows its merged cells span", async () => {
    const view = await caretInTable();

    expect(harnessManager().requireCommand("selectRow")()(view.state, view.dispatch)).toBe(true);

    const selection = view.state.selection as CellSelection;
    expect(selection).toBeInstanceOf(CellSelection);
    expect(selection.isRowSelection()).toBe(true);
    const cells: string[] = [];
    selection.forEachCell((cell) => {
      cells.push(cell.firstChild?.textContent ?? "");
    });
    expect(cells).toEqual(["Cell A2", "Cell B1", "Outer", "Cell B3", "Cell C3"]);
  });
});
