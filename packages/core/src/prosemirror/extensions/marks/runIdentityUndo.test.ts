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
    const caret = before.selection.from;
    const view = new HeadlessEditorView(before);
    view.typeText("x");
    const typed = view.state;
    const typedMarks = (state: EditorState) =>
      (state.doc.nodeAt(caret)?.marks ?? []).map((mark) => mark.type.name);
    expect(typed.doc.textBetween(caret, caret + 1)).toBe("x");
    expect(typedMarks(typed)).not.toContain("runIdentity");

    const undone = runHistory(typed, undo);
    expect(undone.doc.eq(before.doc)).toBe(true);

    const redone = runHistory(undone, redo);
    expect(redone.doc.eq(typed.doc)).toBe(true);
    expect(typedMarks(redone)).not.toContain("runIdentity");
  });
});
