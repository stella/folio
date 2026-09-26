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

// The shape's outer table: four grid columns; rows 1–2 hold a 2×2 merged
// block, a plain column and a vertical merge in column D; rows 3–4 a vertical
// merge in column A and a nested table.
const caretInTable = async () => {
  const document = await parseShapeDocument(await documentShape("tables").build());
  const state = placeSelection(createHarnessState(document, "editing"), "Cell C2", "caret-middle");
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
      "Cell D1",
      "Cell C2",
      "Cell A3",
      "Cell B3",
      "Outer",
      "Cell D3",
      "Cell B4",
      "Cell C4",
      "Cell D4",
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
    expect(cells).toEqual(["Merged A1", "Cell C1", "Cell D1", "Cell C2"]);
  });
});
