/**
 * Edits that change more than text — removing a hyperlink, deleting a note
 * reference, inserting a table — are tracked when suggesting: rejecting every
 * change gives the document back, and accepting every change gives what the
 * same edit gives outside suggesting mode.
 */

import { describe, expect, test } from "bun:test";

import JSZip from "jszip";
import { TextSelection, type EditorState } from "prosemirror-state";

import { documentShape } from "../__tests__/documentShapes";
import { CONFORMANCE_OPERATIONS } from "../__tests__/editorCommandConformance";
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
import { fromProseDoc } from "./conversion/fromProseDoc";
import { createNoteReferenceFollower, withoutUnreferencedNotes } from "./noteReferenceReview";

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

/**
 * Save as the editors do: the notes follow the body's references from
 * `before` to `after`, and a note nothing refers to is dropped.
 */
const saveLikeEditors = async (before: EditorState, after: EditorState, base: Document) => {
  const follower = createNoteReferenceFollower();
  follower.noteBase(before.doc);
  const written = follower.reconcile(fromProseDoc(after.doc, base), after.doc);
  return saveHarnessState(after, withoutUnreferencedNotes(written));
};

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
  const saved = await saveLikeEditors(before, view.state, base);
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
    const accepted = await saveLikeEditors(after, resolveAllChanges(after, "accept"), base);
    expect(normalFootnotes(accepted.model)).toEqual([]);
    const rejected = await saveLikeEditors(after, resolveAllChanges(after, "reject"), base);
    expect(normalFootnotes(rejected.model)).toEqual(normalFootnotes(base));

    // An editor saves again against what its last save produced.
    const acceptedLater = await saveLikeEditors(
      after,
      resolveAllChanges(after, "accept"),
      saved.model,
    );
    expect(normalFootnotes(acceptedLater.model)).toEqual([]);
    const rejectedLater = await saveLikeEditors(
      after,
      resolveAllChanges(after, "reject"),
      saved.model,
    );
    expect(normalFootnotes(rejectedLater.model)).toEqual(normalFootnotes(base));
  });

  test("deleting a note reference outright removes the note", async () => {
    const { saved } = await run(NOTE_CASE, "editing");
    expect(normalFootnotes(saved.model)).toEqual([]);
  });

  test("a note made in the editor is deleted with its own reference mark", async () => {
    const shape = documentShape("notes");
    const base = await parseShapeDocument(await shape.build());
    const caret = placeSelection(createHarnessState(base, "suggesting"), shape.focus, "caret-end");
    if (!caret) {
      throw new Error("No caret");
    }
    const view = new HeadlessEditorView(caret);
    const insertNote = CONFORMANCE_OPERATIONS.find(({ id }) => id === "command:insertFootnote");
    expect(insertNote?.run({ view, base, focus: shape.focus })).toBe(true);
    const added = normalFootnotes(base).at(-1);
    if (!added) {
      throw new Error("No note was added");
    }
    // Accept the insertion (deleting one's own pending insertion retracts it
    // instead), then select the new reference and delete it.
    view.state = resolveAllChanges(view.state, "accept");
    let reference: { from: number; to: number } | null = null;
    view.state.doc.descendants((node, pos) => {
      const mark = node.marks.find(({ type }) => type.name === "footnoteRef");
      if (node.isText && mark && String(mark.attrs["id"]) === String(added.id)) {
        reference = { from: pos, to: pos + node.nodeSize };
      }
    });
    if (!reference) {
      throw new Error("No reference to the new note");
    }
    const { from, to } = reference;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
    const accepted = view.state;
    expect(harnessManager().requireCommand("deleteNoteRef")()(view.state, view.dispatch)).toBe(
      true,
    );

    const { bytes } = await saveLikeEditors(accepted, view.state, base);
    const footnotesXml = await (
      await JSZip.loadAsync(bytes)
    )
      .file("word/footnotes.xml")
      ?.async("text");
    const note = new RegExp(
      `<w:footnote\\b[^>]*\\bw:id="${added.id}"[^>]*>[\\s\\S]*?</w:footnote>`,
      "u",
    ).exec(footnotesXml ?? "")?.[0];
    expect(note).toBeDefined();
    // The only reference mark is the one inside a deletion.
    expect(note?.match(/<w:footnoteRef\/>/gu)).toHaveLength(1);
    expect(note).toMatch(/<w:del\b[^>]*>(?:(?!<\/w:del>)[\s\S])*<w:footnoteRef\/>/u);
  });

  test("a note whose pending reference is taken back is not saved", async () => {
    const shape = documentShape("notes");
    const base = await parseShapeDocument(await shape.build());
    const caret = placeSelection(createHarnessState(base, "suggesting"), shape.focus, "caret-end");
    if (!caret) {
      throw new Error("No caret");
    }
    const view = new HeadlessEditorView(caret);
    const insertNote = CONFORMANCE_OPERATIONS.find(({ id }) => id === "command:insertFootnote");
    expect(insertNote?.run({ view, base, focus: shape.focus })).toBe(true);
    const added = normalFootnotes(base).at(-1);
    if (!added) {
      throw new Error("No note was added");
    }
    // The reference is the author's own pending insertion: deleting it takes it back.
    const inserted = view.state;
    const { from } = view.state.selection;
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, from - String(added.id).length, from),
      ),
    );
    expect(harnessManager().requireCommand("deleteNoteRef")()(view.state, view.dispatch)).toBe(
      true,
    );

    const { model } = await saveLikeEditors(inserted, view.state, base);
    expect(normalFootnotes(model).map(({ id }) => id)).not.toContain(added.id);
    expect(normalFootnotes(model).length).toBe(normalFootnotes(base).length - 1);
  });
});
