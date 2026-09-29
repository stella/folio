/**
 * Edits that change more than text — removing a hyperlink, deleting a note
 * reference, inserting a table — are tracked when suggesting: rejecting every
 * change gives the document back, and accepting every change gives what the
 * same edit gives outside suggesting mode.
 */

import { describe, expect, test } from "bun:test";

import { documentShape } from "../__tests__/documentShapes";
import {
  createHarnessState,
  type EditorMode,
  harnessManager,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  resolveAllChanges,
  saveHarnessState,
  type SelectionPlacement,
  summarizeState,
} from "../__tests__/editorHarness";

type Case = {
  shape: string;
  placement: SelectionPlacement;
  command: string;
  args: readonly unknown[];
};

const CASES: readonly Case[] = [
  { shape: "fields-links-bookmarks", placement: "paragraph", command: "removeHyperlink", args: [] },
  { shape: "notes", placement: "paragraph", command: "deleteNoteRef", args: [] },
  { shape: "plain-markdown", placement: "caret-middle", command: "insertTable", args: [2, 2] },
];

const run = async ({ shape: shapeId, placement, command, args }: Case, mode: EditorMode) => {
  const shape = documentShape(shapeId);
  const base = await parseShapeDocument(await shape.build());
  const before = placeSelection(createHarnessState(base, mode), shape.focus, placement);
  if (!before) {
    throw new Error("No selection");
  }
  const view = new HeadlessEditorView(before);
  expect(
    harnessManager().requireCommand(command)(...args)(view.state, view.dispatch, view as never),
  ).toBe(true);
  expect(view.state.doc.eq(before.doc)).toBe(false);
  await saveHarnessState(view.state, base);
  return { before, after: view.state };
};

describe("structural edits in suggesting mode", () => {
  test.each(CASES.map((entry) => [entry.command, entry] as const))(
    "%s is tracked: rejecting restores, accepting matches editing",
    async (_command, entry) => {
      const editing = await run(entry, "editing");
      const suggesting = await run(entry, "suggesting");
      expect(summarizeState(resolveAllChanges(suggesting.after, "reject"))).toEqual(
        summarizeState(resolveAllChanges(suggesting.before, "reject")),
      );
      expect(summarizeState(resolveAllChanges(suggesting.after, "accept"))).toEqual(
        summarizeState(resolveAllChanges(editing.after, "accept")),
      );
    },
  );
});
