/**
 * Text typed over a selection that ends between blocks: selected cells, whose
 * range spans the block boundaries inside the cell, or a selected block.
 */

import { describe, expect, test } from "bun:test";
import { type EditorState, NodeSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { shapeArrayBuffer } from "./documentShapes";
import {
  createHarnessState,
  type EditorMode,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
} from "./editorHarness";

const tablesState = async (mode: EditorMode) =>
  createHarnessState(
    await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("tables"))),
    mode,
  );

const cellPositions = (state: EditorState) => {
  const cells: number[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name === "tableCell") cells.push(pos);
    return true;
  });
  return cells;
};

const cellTexts = (state: EditorState) =>
  cellPositions(state).map((pos) => {
    const texts: string[] = [];
    state.doc.nodeAt(pos)?.forEach((child) => {
      texts.push(child.textContent);
    });
    return texts;
  });

const typedOverCells = async (mode: EditorMode, deletesFirst: boolean) => {
  const state = await tablesState(mode);
  const [anchor, head] = cellPositions(state);
  if (anchor === undefined || head === undefined)
    throw new Error("Fixture has no adjacent cells");
  const view = new HeadlessEditorView(
    state.apply(
      state.tr.setSelection(CellSelection.create(state.doc, anchor, head)),
    ),
  );
  if (deletesFirst) view.pressKey("Delete");
  view.typeText("alpha");
  return { initial: state, typed: view.state };
};

const caretFollowsTypedText = (state: EditorState) => {
  const { selection } = state;
  expect(selection.empty).toBe(true);
  expect(
    selection.$head.parent.textContent.slice(0, selection.$head.parentOffset),
  ).toEndWith("alpha");
};

describe("typing over selected cells", () => {
  for (const deletesFirst of [true, false]) {
    const when = deletesFirst ? "after Delete" : "directly";

    test(`outside suggesting mode keeps the text in order, ${when}`, async () => {
      const { typed } = await typedOverCells("editing", deletesFirst);
      expect(cellTexts(typed).flat()).toContain("alpha");
      caretFollowsTypedText(typed);
    });

    test(`in suggesting mode tracks the same edit, ${when}`, async () => {
      const { initial, typed } = await typedOverCells(
        "suggesting",
        deletesFirst,
      );
      const direct = await typedOverCells("editing", deletesFirst);
      caretFollowsTypedText(typed);
      expect(cellTexts(resolveAllChanges(typed, "accept"))).toEqual(
        cellTexts(direct.typed),
      );
      expect(cellTexts(resolveAllChanges(typed, "reject"))).toEqual(
        cellTexts(initial),
      );
    });
  }
});

describe("typing over a selected block", () => {
  for (const mode of ["editing", "suggesting"] as const) {
    test(`keeps the text in order with the caret after it, ${mode}`, async () => {
      const state = await tablesState(mode);
      const rule = state.schema.nodes["horizontalRule"];
      if (!rule) throw new Error("Schema has no horizontal rule");
      const withRule = state.apply(state.tr.insert(0, rule.create()));
      const view = new HeadlessEditorView(
        withRule.apply(
          withRule.tr.setSelection(NodeSelection.create(withRule.doc, 0)),
        ),
      );
      view.typeText("alpha");

      const { selection } = view.state;
      expect(selection.$head.parent.textContent).toBe("alpha");
      caretFollowsTypedText(view.state);
      if (mode === "suggesting") {
        expect(
          selection.$head.parent.firstChild?.marks.map(({ type }) => type.name),
        ).toContain("insertion");
      }
    });
  }
});
