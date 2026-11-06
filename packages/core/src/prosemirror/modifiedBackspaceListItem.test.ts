/**
 * Ctrl, Cmd or Shift with Backspace at the start of a list item joins it to the
 * paragraph before. In suggesting mode that join is recorded: rejecting gives
 * the two paragraphs back, accepting gives what the direct edit gives.
 */

import { expect, test } from "bun:test";

import { documentShape } from "../__tests__/documentShapes";
import {
  createHarnessState,
  type EditorMode,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  resolveAllChanges,
  summarizeState,
} from "../__tests__/editorHarness";

const pressed = async (shapeId: string, key: string, mode: EditorMode) => {
  const shape = documentShape(shapeId);
  const base = await parseShapeDocument(await shape.build());
  const before = placeSelection(
    createHarnessState(base, mode),
    shape.focus,
    "caret-start",
  );
  if (!before) throw new Error("No selection");
  const view = new HeadlessEditorView(before);
  view.pressKey(key);
  return { before, after: view.state };
};

test.each(
  [
    "single-decimal-list",
    "single-bullet-list",
    "outline-level-numbered",
  ].flatMap((shape) =>
    ["Mod-Backspace", "Shift-Backspace"].map((key) => [shape, key]),
  ),
)(
  "%s, %s: rejecting restores, accepting matches editing",
  async (shape, key) => {
    const suggesting = await pressed(
      shape as string,
      key as string,
      "suggesting",
    );
    const direct = await pressed(shape as string, key as string, "editing");
    expect(direct.after.doc.eq(direct.before.doc)).toBe(false);
    expect(
      summarizeState(resolveAllChanges(suggesting.after, "reject")),
    ).toEqual(summarizeState(resolveAllChanges(suggesting.before, "reject")));
    expect(
      summarizeState(resolveAllChanges(suggesting.after, "accept")),
    ).toEqual(summarizeState(resolveAllChanges(direct.after, "accept")));
  },
);
