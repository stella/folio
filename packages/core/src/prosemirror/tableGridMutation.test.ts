import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { documentShape } from "../__tests__/documentShapes";
import {
  createHarnessState,
  harnessManager,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  saveHarnessState,
} from "../__tests__/editorHarness";

/**
 * The shape's table follows a paragraph: row 1 holds a 2×2 merge (A1:B2),
 * "Cell C1" and a vertical merge D1:D2, so row 2 holds only "Cell C2"; A3:A4
 * is merged too. The harness binds its commands to an editor schema of its
 * own, as a header or footer editor does.
 */
const outerTable = (doc: PMNode): PMNode => {
  let found: PMNode | null = null;
  doc.descendants((node) => {
    if (!found && node.type.name === "table") {
      found = node;
    }
    return !found;
  });
  if (!found) {
    throw new Error("No table");
  }
  return found;
};

/** Each row's cells as `text@column rowspan×colspan`, after checking the grid tiles. */
const layout = (state: EditorState): string[][] => {
  const table = outerTable(state.doc);
  const map = TableMap.get(table);
  expect(map.problems ?? []).toEqual([]);
  expect(table.attrs["columnWidths"]).toHaveLength(map.width);
  const rows: string[][] = [];
  let rowStart = 0;
  table.forEach((row) => {
    const cells: string[] = [];
    let cellStart = rowStart + 1;
    row.forEach((cell) => {
      cells.push(
        `${cell.textContent}@${map.findCell(cellStart).left} ${cell.attrs["rowspan"]}×${cell.attrs["colspan"]}`,
      );
      cellStart += cell.nodeSize;
    });
    rows.push(cells);
    rowStart += row.nodeSize;
  });
  return rows;
};

const runAtC2 = async (command: string) => {
  const base = await parseShapeDocument(await documentShape("tables").build());
  const state = placeSelection(createHarnessState(base, "editing"), "Cell C2", "caret-middle");
  if (!state) {
    throw new Error("No caret");
  }
  const view = new HeadlessEditorView(state);
  expect(harnessManager().requireCommand(command)()(view.state, view.dispatch, view as never)).toBe(
    true,
  );
  await saveHarnessState(view.state, base);
  return view.state;
};

describe("row and column commands beside vertical merges", () => {
  test("adding a row above grows the merges the new row passes through", async () => {
    expect(layout(await runAtC2("addRowAbove"))).toEqual([
      ["Merged A1@0 3×2", "Cell C1@2 1×1", "Cell D1@3 3×1"],
      ["@2 1×1"],
      ["Cell C2@2 1×1"],
      ["Cell A3@0 2×1", "Cell B3@1 1×1", "OuterInner 1Inner 2@2 1×1", "Cell D3@3 1×1"],
      ["Cell B4@1 1×1", "Cell C4@2 1×1", "Cell D4@3 1×1"],
    ]);
  });

  test("adding a column left puts each row's new cell in the grid column", async () => {
    const state = await runAtC2("addColumnLeft");
    expect(layout(state)).toEqual([
      ["Merged A1@0 2×2", "@2 1×1", "Cell C1@3 1×1", "Cell D1@4 2×1"],
      ["@2 1×1", "Cell C2@3 1×1"],
      ["Cell A3@0 2×1", "Cell B3@1 1×1", "@2 1×1", "OuterInner 1Inner 2@3 1×1", "Cell D3@4 1×1"],
      ["Cell B4@1 1×1", "@2 1×1", "Cell C4@3 1×1", "Cell D4@4 1×1"],
    ]);
    expect(outerTable(state.doc).attrs["columnWidths"]).toEqual([2200, 2200, 2200, 2200, 2200]);
  });

  test("adding a column right puts each row's new cell in the grid column", async () => {
    expect(layout(await runAtC2("addColumnRight"))).toEqual([
      ["Merged A1@0 2×2", "Cell C1@2 1×1", "@3 1×1", "Cell D1@4 2×1"],
      ["Cell C2@2 1×1", "@3 1×1"],
      ["Cell A3@0 2×1", "Cell B3@1 1×1", "OuterInner 1Inner 2@2 1×1", "@3 1×1", "Cell D3@4 1×1"],
      ["Cell B4@1 1×1", "Cell C4@2 1×1", "@3 1×1", "Cell D4@4 1×1"],
    ]);
  });

  test("deleting a column removes each row's cell in it and keeps the grid", async () => {
    const state = await runAtC2("deleteColumn");
    // Row 2 had no cell outside the merges, so it goes and the merges shorten.
    expect(layout(state)).toEqual([
      ["Merged A1@0 1×2", "Cell D1@2 1×1"],
      ["Cell A3@0 2×1", "Cell B3@1 1×1", "Cell D3@2 1×1"],
      ["Cell B4@1 1×1", "Cell D4@2 1×1"],
    ]);
    expect(outerTable(state.doc).attrs["columnWidths"]).toEqual([2200, 2200, 2200]);
  });
});
