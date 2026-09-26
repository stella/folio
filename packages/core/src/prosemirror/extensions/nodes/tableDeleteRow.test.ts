import { describe, expect, test } from "bun:test";
import { TableMap } from "prosemirror-tables";
import type { Node as PMNode } from "prosemirror-model";

import { documentShape } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  harnessManager,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
} from "../../../__tests__/editorHarness";

const outerTable = (doc: PMNode): PMNode => {
  const table = doc.children.find((node) => node.type.spec["tableRole"] === "table");
  if (!table) {
    throw new Error("the document holds no table");
  }
  return table;
};

const rowTexts = (table: PMNode): string[][] =>
  table.children.map((row) =>
    row.children.map(
      (cell) => `${cell.firstChild?.textContent ?? ""}/${String(cell.attrs["rowspan"])}`,
    ),
  );

describe("deleteRow beside merged cells", () => {
  test("a vertical merge right of a wide merged cell closes over the removed row", async () => {
    // Row 2 holds only the plain cell C2: a 2×2 block covers columns A–B and
    // a vertical merge covers column D, both from row 1.
    const document = await parseShapeDocument(await documentShape("tables").build());
    const before = placeSelection(
      createHarnessState(document, "editing"),
      "Cell C2",
      "caret-middle",
    );
    if (!before) {
      throw new Error("the shape has no focus paragraph");
    }
    const view = new HeadlessEditorView(before);

    expect(harnessManager().requireCommand("deleteRow")()(view.state, view.dispatch)).toBe(true);

    const table = outerTable(view.state.doc);
    const map = TableMap.get(table);
    expect(map.problems ?? []).toEqual([]);
    expect(map.width).toBe(4);
    expect(rowTexts(table)[0]).toEqual(["Merged A1/1", "Cell C1/1", "Cell D1/1"]);
  });
});
