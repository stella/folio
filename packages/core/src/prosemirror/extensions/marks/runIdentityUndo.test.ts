import { describe, expect, test } from "bun:test";
import { redo, undo } from "prosemirror-history";
import type { EditorState } from "prosemirror-state";

import { documentShape } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
} from "../../../__tests__/editorHarness";

const runHistory = (state: EditorState, command: typeof undo): EditorState => {
  let next = state;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
};

const runIdentityIds = (state: EditorState): unknown[] => {
  const ids: unknown[] = [];
  state.doc.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name === "runIdentity") {
        ids.push(mark.attrs["id"]);
      }
    }
    return true;
  });
  return ids;
};

describe("run identity through undo", () => {
  test("undoing a deletion restores the deleted text's run identity", async () => {
    const document = await parseShapeDocument(await documentShape("plain-markdown").build());
    const before = placeSelection(createHarnessState(document, "editing"), "First item", "word");
    if (!before) {
      throw new Error("the shape has no focus paragraph");
    }
    expect(runIdentityIds(before).length).toBeGreaterThan(0);

    const view = new HeadlessEditorView(before);
    view.pressKey("Backspace");
    expect(view.state.doc.eq(before.doc)).toBe(false);

    expect(runHistory(view.state, undo).doc.eq(before.doc)).toBe(true);
  });

  test("redoing typed text does not give it the run it was typed into", async () => {
    const document = await parseShapeDocument(await documentShape("plain-markdown").build());
    const before = placeSelection(
      createHarnessState(document, "editing"),
      "First item",
      "caret-middle",
    );
    if (!before) {
      throw new Error("the shape has no focus paragraph");
    }
    const view = new HeadlessEditorView(before);
    view.typeText("x");
    const typed = view.state;
    const redone = runHistory(runHistory(typed, undo), redo);

    expect(redone.doc.eq(typed.doc)).toBe(true);
  });
});
