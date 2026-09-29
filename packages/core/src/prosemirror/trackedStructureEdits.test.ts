/**
 * Edits that change more than text — removing a hyperlink, deleting a note
 * reference, inserting a table — are tracked when suggesting: rejecting every
 * change gives the document back, and accepting every change gives what the
 * same edit gives outside suggesting mode.
 */

import { describe, expect, test } from "bun:test";

import type { EditorState } from "prosemirror-state";

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
import type { Document } from "../types/document";

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
  const saved = await saveHarnessState(view.state, base);
  return { base, before, after: view.state, saved };
};

const HYPERLINK_CASE = CASES[0] as Case;
const NOTE_CASE = CASES[1] as Case;

/** The marks of every text node reading `text`, as type names with the run-property change's previous style. */
const textRuns = (state: EditorState, text: string) => {
  const runs: { marks: string[]; previousStyle: unknown }[] = [];
  state.doc.descendants((node) => {
    if (node.isText && node.text === text) {
      const change = node.marks.find(({ type }) => type.name === "runPropertyChange");
      const changes: unknown = change?.attrs["changes"];
      runs.push({
        marks: node.marks.map(({ type }) => type.name).toSorted(),
        previousStyle: Array.isArray(changes)
          ? (changes[0] as { previousFormatting?: { styleId?: unknown } }).previousFormatting
              ?.styleId
          : undefined,
      });
    }
  });
  return runs;
};

/** The saved model's normal footnotes. */
const normalFootnotes = (model: Document) =>
  (model.package.footnotes ?? []).filter(
    (note) => note.noteType === undefined || note.noteType === "normal",
  );

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

  test("a tracked hyperlink removal records the Hyperlink style on both runs", async () => {
    const { after } = await run(HYPERLINK_CASE, "suggesting");
    const [deleted, inserted] = textRuns(after, "site");
    // The struck text keeps its link but no longer states the style; the
    // text that follows never had either.
    expect(deleted?.marks).toContain("hyperlink");
    expect(deleted?.marks).toContain("deletion");
    expect(inserted?.marks).toContain("insertion");
    expect(inserted?.marks).not.toContain("hyperlink");
    for (const textRun of [deleted, inserted]) {
      expect(textRun?.marks).not.toContain("characterStyle");
      expect(textRun?.previousStyle).toBe("Hyperlink");
    }

    const [restored] = textRuns(resolveAllChanges(after, "reject"), "site");
    expect(restored?.marks).toContain("hyperlink");
    expect(restored?.marks).toContain("characterStyle");
    expect(restored?.marks).not.toContain("runPropertyChange");
  });

  test("removing a hyperlink drops its Hyperlink style", async () => {
    const { after } = await run(HYPERLINK_CASE, "editing");
    const marks: string[] = [];
    after.doc.descendants((node) => {
      if (node.isText && node.text?.includes("site")) {
        marks.push(...node.marks.map(({ type }) => type.name));
      }
    });
    expect(marks).not.toContain("hyperlink");
    expect(marks).not.toContain("characterStyle");
  });

  test("a tracked note reference deletion deletes the note's content with it", async () => {
    const { base, after, saved } = await run(NOTE_CASE, "suggesting");
    const [note] = normalFootnotes(saved.model);
    expect(note).toBeDefined();
    for (const block of note?.content ?? []) {
      if (block.type !== "paragraph") {
        continue;
      }
      expect(block.pPrMark?.kind).toBe("del");
      for (const item of block.content) {
        expect(item.type === "run").toBe(false);
      }
    }

    // Accepting removes the note; rejecting keeps it as it was.
    const accepted = await saveHarnessState(resolveAllChanges(after, "accept"), base);
    expect(normalFootnotes(accepted.model)).toEqual([]);
    const rejected = await saveHarnessState(resolveAllChanges(after, "reject"), base);
    expect(normalFootnotes(rejected.model)).toEqual(normalFootnotes(base));

    // An editor saves again against what its last save produced.
    const acceptedLater = await saveHarnessState(resolveAllChanges(after, "accept"), saved.model);
    expect(normalFootnotes(acceptedLater.model)).toEqual([]);
    const rejectedLater = await saveHarnessState(resolveAllChanges(after, "reject"), saved.model);
    expect(normalFootnotes(rejectedLater.model)).toEqual(normalFootnotes(base));
  });

  test("deleting a note reference outright removes the note", async () => {
    const { saved } = await run(NOTE_CASE, "editing");
    expect(normalFootnotes(saved.model)).toEqual([]);
  });
});
