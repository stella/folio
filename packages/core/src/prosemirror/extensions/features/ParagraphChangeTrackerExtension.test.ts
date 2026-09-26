/**
 * Unit tests for ParagraphChangeTrackerExtension
 */

import { describe, test, expect } from "bun:test";
import { Schema, Slice } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from "prosemirror-transform";

import {
  getChangedParagraphIds,
  hasStructuralChanges,
  hasUntrackedChanges,
  clearTrackedChanges,
  ignoreTrackedChanges,
  markChangedParagraphRanges,
  ParagraphChangeTrackerExtension,
} from "./ParagraphChangeTrackerExtension";

import { ParaIdAllocatorExtension } from "./ParaIdAllocatorExtension";

// Minimal schema with paraId support
const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: {
        paraId: { default: null },
        textId: { default: null },
        pPrMark: { default: null },
      },
      toDOM: () => ["p", 0],
    },
    table: {
      group: "block",
      content: "tableRow+",
      attrs: { width: { default: null } },
    },
    tableRow: { content: "tableCell+" },
    tableCell: { content: "paragraph+" },
    text: { group: "inline" },
  },
  marks: {
    bold: {
      parseDOM: [{ tag: "strong" }],
      toDOM: () => ["strong", 0],
    },
  },
});

// Get the plugin from the extension
const ext = ParagraphChangeTrackerExtension();
const runtime = ext.onSchemaReady({ schema });
const plugin = runtime.plugins?.[0];
if (!plugin) {
  throw new Error("Expected plugin from ParagraphChangeTrackerExtension");
}

function createDoc(...paras: { text: string; paraId?: string }[]) {
  return schema.node(
    "doc",
    null,
    paras.map((p) =>
      schema.node("paragraph", { paraId: p.paraId ?? null }, p.text ? [schema.text(p.text)] : []),
    ),
  );
}

function createState(paras: { text: string; paraId?: string }[]) {
  const doc = createDoc(...paras);
  return EditorState.create({ doc, plugins: [plugin] });
}

function typeText(state: EditorState, text: string, pos?: number): EditorState {
  const insertPos = pos ?? state.selection.from;
  const tr = state.tr.insertText(text, insertPos);
  return state.apply(tr);
}

function deleteRange(state: EditorState, from: number, to: number): EditorState {
  const tr = state.tr.delete(from, to);
  return state.apply(tr);
}

function setSelection(state: EditorState, pos: number): EditorState {
  const tr = state.tr.setSelection(TextSelection.create(state.doc, pos));
  return state.apply(tr);
}

// ============================================================================
// Tests
// ============================================================================

describe("ParagraphChangeTrackerExtension", () => {
  describe("mark-only edits", () => {
    test("does not crash when a mark step is followed by a shrinking replace step", () => {
      let state = createState([
        { text: "AAAA", paraId: "P1" },
        { text: "BBBB", paraId: "P2" },
      ]);
      const bold = schema.marks.bold;
      const boldMark = bold.create();

      state = state.apply(state.tr.step(new AddMarkStep(7, 11, boldMark)));

      const tr = state.tr;
      tr.step(new RemoveMarkStep(7, 11, boldMark));
      tr.step(new ReplaceStep(1, 5, Slice.empty));

      expect(() => {
        state = state.apply(tr);
      }).not.toThrow();

      const changed = getChangedParagraphIds(state);
      expect(changed.has("P1")).toBe(true);
      expect(changed.has("P2")).toBe(true);
    });
  });

  describe("single paragraph edit", () => {
    test("tracks changed paraId when text is inserted", () => {
      let state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      // Type in the first paragraph (position 1 = inside first para)
      state = typeText(state, " there", 6); // After "Hello"

      const changed = getChangedParagraphIds(state);
      expect(changed.has("P1")).toBe(true);
      expect(changed.has("P2")).toBe(false);
    });

    test("tracks changed paraId when text is deleted", () => {
      let state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      // Delete "lo" from "Hello" (positions 4-6 in doc)
      state = deleteRange(state, 4, 6);

      const changed = getChangedParagraphIds(state);
      expect(changed.has("P1")).toBe(true);
      expect(changed.has("P2")).toBe(false);
    });
  });

  describe("multi-paragraph formatting", () => {
    test("tracks multiple paraIds when editing different paragraphs", () => {
      let state = createState([
        { text: "First", paraId: "P1" },
        { text: "Second", paraId: "P2" },
        { text: "Third", paraId: "P3" },
      ]);

      // Insert inside P1 (pos 2 = inside first paragraph)
      state = typeText(state, "X", 2);
      expect(getChangedParagraphIds(state).has("P1")).toBe(true);

      // Find P3 start position dynamically and insert there
      let p3Start = 0;
      state.doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.attrs.paraId === "P3") {
          p3Start = pos + 1; // Inside the paragraph
        }
      });
      state = typeText(state, "Y", p3Start);

      const changed = getChangedParagraphIds(state);
      expect(changed.has("P1")).toBe(true);
      expect(changed.has("P3")).toBe(true);
    });
  });

  describe("structural changes", () => {
    test("detects paragraph split (Enter key creates new paragraph)", () => {
      let state = createState([{ text: "Hello World", paraId: "P1" }]);

      // Split the paragraph: replace text from pos 6 to 6 with a new paragraph node
      const tr = state.tr.split(6);
      state = state.apply(tr);

      expect(hasStructuralChanges(state)).toBe(true);
    });

    test("detects paragraph merge (join)", () => {
      let state = createState([
        { text: "First", paraId: "P1" },
        { text: "Second", paraId: "P2" },
      ]);

      // Join at the boundary between the two paragraphs
      // End of P1 is at position 6, start of P2 is at position 7
      const tr = state.tr.join(7);
      state = state.apply(tr);

      expect(hasStructuralChanges(state)).toBe(true);
    });
  });

  describe("no-edit scenario", () => {
    test("has empty changed set when no edits are made", () => {
      const state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      expect(getChangedParagraphIds(state).size).toBe(0);
      expect(hasStructuralChanges(state)).toBe(false);
    });

    test("has empty changed set after selection-only change", () => {
      let state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      // Just move the cursor — no content change
      state = setSelection(state, 3);

      expect(getChangedParagraphIds(state).size).toBe(0);
      expect(hasStructuralChanges(state)).toBe(false);
    });
  });

  describe("paragraphs without paraId", () => {
    test("sets hasUntrackedChanges when editing paragraph with no paraId", () => {
      let state = createState([
        { text: "Hello", paraId: undefined },
        { text: "World", paraId: "P2" },
      ]);

      // Edit the first paragraph which has no paraId
      state = typeText(state, "X", 1);

      expect(hasUntrackedChanges(state)).toBe(true);
    });

    test("does not set hasUntrackedChanges when editing tracked paragraphs", () => {
      let state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      state = typeText(state, "X", 1);
      expect(hasUntrackedChanges(state)).toBe(false);
    });
  });

  describe("clear after save", () => {
    test("clears all tracked state", () => {
      let state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      // Make some edits
      state = typeText(state, "X", 1);

      expect(getChangedParagraphIds(state).size).toBeGreaterThan(0);

      // Clear tracked changes
      const clearTr = clearTrackedChanges(state);
      state = state.apply(clearTr);

      expect(getChangedParagraphIds(state).size).toBe(0);
      expect(hasStructuralChanges(state)).toBe(false);
      expect(hasUntrackedChanges(state)).toBe(false);
    });

    test("tracks new changes after clear", () => {
      let state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      // Edit P1
      state = typeText(state, "X", 1);

      // Clear
      state = state.apply(clearTrackedChanges(state));

      // Edit P2 (position after P1: doc[0]=p1(6 chars), doc[1]=p2 starts at 8)
      state = typeText(state, "Y", 9);

      const changed = getChangedParagraphIds(state);
      expect(changed.has("P1")).toBe(false);
      expect(changed.has("P2")).toBe(true);
    });
  });

  describe("ignored generated transactions", () => {
    test("does not record ignored doc-changing transactions", () => {
      let state = createState([{ text: "Hello", paraId: "P1" }]);

      const tr = state.tr.setNodeMarkup(0, undefined, {
        ...state.doc.child(0).attrs,
        paraId: "P2",
      });
      state = state.apply(ignoreTrackedChanges(tr));

      expect(getChangedParagraphIds(state).size).toBe(0);
      expect(hasStructuralChanges(state)).toBe(false);
      expect(hasUntrackedChanges(state)).toBe(false);
    });

    test("preserves existing tracked edits while ignoring generated changes", () => {
      let state = createState([
        { text: "Hello", paraId: "P1" },
        { text: "World", paraId: "P2" },
      ]);

      state = typeText(state, "X", 1);
      expect(getChangedParagraphIds(state).has("P1")).toBe(true);

      let secondParagraphPos = 0;
      state.doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.attrs.paraId === "P2") {
          secondParagraphPos = pos;
          return false;
        }
        return true;
      });

      const tr = state.tr.setNodeMarkup(secondParagraphPos, undefined, {
        ...state.doc.child(1).attrs,
        paraId: "P3",
      });
      state = state.apply(ignoreTrackedChanges(tr));

      const changed = getChangedParagraphIds(state);
      expect(changed.has("P1")).toBe(true);
      expect(changed.has("P2")).toBe(false);
      expect(changed.has("P3")).toBe(false);
    });
  });

  describe("attribute-only edits", () => {
    test("tracks a paragraph attribute step mapped through a later join", () => {
      const state = createState([
        { text: "before", paraId: "P1" },
        { text: "middle", paraId: "P2" },
        { text: "after", paraId: "P3" },
      ]);
      const secondPos = state.doc.child(0).nodeSize;
      const thirdPos = secondPos + state.doc.child(1).nodeSize;
      const tr = state.tr;
      tr.setNodeAttribute(thirdPos, "pPrMark", null);
      tr.setNodeAttribute(secondPos, "pPrMark", null);
      tr.join(thirdPos);

      const next = state.apply(tr);
      expect(next.doc.childCount).toBe(2);
      expect(getChangedParagraphIds(next).has("P2")).toBe(true);
      expect(hasStructuralChanges(next)).toBe(true);
    });

    // A blank paragraph has no text to mark, so its paragraph mark, its list
    // level and its style are the only things about it that CAN change. Those
    // arrive as `AttrStep`s, whose step map is empty: read the position off
    // the step, or the save writes the paragraph's original XML and the edit
    // is gone from the file while the editor still shows it.
    test("tracks the paragraph an attribute step changed", () => {
      const state = createState([
        { text: "kept", paraId: "P1" },
        { text: "", paraId: "P2" },
      ]);
      let position = 0;
      state.doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.attrs["paraId"] === "P2") {
          position = pos;
        }
      });

      const next = state.apply(
        state.tr.setNodeAttribute(position, "pPrMark", {
          kind: "del",
          info: { id: 1, author: "a", date: "2000-01-01T00:00:00.000Z" },
        }),
      );

      expect(getChangedParagraphIds(next).has("P2")).toBe(true);
      expect(getChangedParagraphIds(next).has("P1")).toBe(false);
      expect(hasStructuralChanges(next)).toBe(false);
    });

    test("reports an attribute step on a paragraph with no paraId as untracked", () => {
      const state = createState([{ text: "" }]);

      const next = state.apply(state.tr.setNodeAttribute(0, "pPrMark", null));

      expect(hasUntrackedChanges(next)).toBe(true);
    });
  });

  test("precise paragraph ranges replace a whole-document step map", () => {
    const state = createState([
      { text: "first", paraId: "P1" },
      { text: "old", paraId: "P2" },
      { text: "last", paraId: "P3" },
    ]);
    const replacement = createDoc(
      { text: "first", paraId: "P1" },
      { text: "new", paraId: "P2" },
      { text: "last", paraId: "P3" },
    );
    const tr = state.tr.replaceWith(0, state.doc.content.size, replacement.content);
    const secondFrom = replacement.child(0).nodeSize;
    markChangedParagraphRanges(tr, {
      ranges: [{ from: secondFrom, to: secondFrom + replacement.child(1).nodeSize }],
      mappingFrom: tr.steps.length,
      replacesStepMapAt: 0,
    });

    const next = state.apply(tr);
    expect([...getChangedParagraphIds(next)]).toEqual(["P2"]);
    expect(hasStructuralChanges(next)).toBe(false);
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  describe("accumulation across multiple transactions", () => {
    test("accumulates changes across multiple edits", () => {
      let state = createState([
        { text: "A", paraId: "P1" },
        { text: "B", paraId: "P2" },
        { text: "C", paraId: "P3" },
      ]);

      // Edit P1
      state = typeText(state, "X", 2);
      expect(getChangedParagraphIds(state).has("P1")).toBe(true);

      // Find P3 position dynamically
      let p3Start = 0;
      state.doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.attrs.paraId === "P3") {
          p3Start = pos + 1;
        }
      });
      state = typeText(state, "Y", p3Start);

      const changed = getChangedParagraphIds(state);
      expect(changed.has("P1")).toBe(true);
      expect(changed.has("P3")).toBe(true);
      expect(changed.has("P2")).toBe(false);
    });
  });
});

describe("paragraph tracking with the ID allocator", () => {
  const allocator = ParaIdAllocatorExtension().onSchemaReady({ schema }).plugins?.at(0);
  if (!allocator) {
    throw new Error("Expected paragraph ID allocator plugin");
  }
  const trackedState = (...paras: { text: string; paraId?: string }[]) =>
    EditorState.create({ doc: createDoc(...paras), plugins: [plugin, allocator] });

  test.each([0, 2, 5])("tracks both split halves at offset %i", (offset) => {
    const initial = trackedState({ text: "Hello", paraId: "11111111" });
    const next = initial.apply(initial.tr.split(offset + 1));
    const firstId = next.doc.child(0).attrs.paraId;
    const secondId = next.doc.child(1).attrs.paraId;

    expect(firstId).toBe("11111111");
    expect(secondId).toMatch(/^[0-9A-F]{8}$/u);
    expect(secondId).not.toBe(firstId);
    expect(getChangedParagraphIds(next)).toEqual(new Set([firstId, secondId]));
    expect(hasStructuralChanges(next)).toBe(true);
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test.each([null, "", "00000000", "11111111", "22222222"])(
    "tracks only the inserted paragraph after allocating its supplied ID %j",
    (paraId) => {
      const initial = trackedState({ text: "Original", paraId: "11111111" });
      const inserted = schema.node("paragraph", { paraId }, schema.text("Inserted"));
      const next = initial.apply(initial.tr.insert(0, inserted));
      const insertedId = next.doc.child(0).attrs.paraId;

      expect(insertedId).toMatch(/^[0-9A-F]{8}$/u);
      expect(insertedId).not.toBe("00000000");
      expect(insertedId).not.toBe("11111111");
      expect(next.doc.child(1).attrs.paraId).toBe("11111111");
      expect(getChangedParagraphIds(next)).toEqual(new Set([insertedId]));
      expect(hasStructuralChanges(next)).toBe(true);
      expect(hasUntrackedChanges(next)).toBe(false);
    },
  );

  test("retains existing edits when a later insertion receives an ID", () => {
    const initial = trackedState(
      { text: "Edited", paraId: "11111111" },
      { text: "Unchanged", paraId: "22222222" },
    );
    const edited = initial.apply(initial.tr.insertText("!", 1));
    const next = edited.apply(edited.tr.insert(0, schema.node("paragraph")));

    expect(getChangedParagraphIds(next)).toEqual(
      new Set(["11111111", next.doc.child(0).attrs.paraId]),
    );
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test("detects deletion without dirtying surviving paragraphs", () => {
    const initial = trackedState(
      { text: "Deleted", paraId: "11111111" },
      { text: "Kept", paraId: "22222222" },
    );
    const next = initial.apply(initial.tr.delete(0, initial.doc.child(0).nodeSize));

    expect(next.doc.childCount).toBe(1);
    expect(getChangedParagraphIds(next)).toEqual(new Set());
    expect(hasStructuralChanges(next)).toBe(true);
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test("tracks the surviving paragraph after a join", () => {
    const initial = trackedState(
      { text: "First", paraId: "11111111" },
      { text: "Second", paraId: "22222222" },
    );
    const next = initial.apply(initial.tr.join(initial.doc.child(0).nodeSize));

    expect(next.doc.child(0).textContent).toBe("FirstSecond");
    expect(getChangedParagraphIds(next)).toEqual(new Set(["11111111"]));
    expect(hasStructuralChanges(next)).toBe(true);
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test("detects count-neutral deletion and insertion in one transaction", () => {
    const initial = trackedState(
      { text: "Replaced", paraId: "11111111" },
      { text: "Kept", paraId: "22222222" },
    );
    const next = initial.apply(
      initial.tr.replaceWith(0, initial.doc.child(0).nodeSize, schema.node("paragraph")),
    );

    expect(next.doc.childCount).toBe(initial.doc.childCount);
    expect(getChangedParagraphIds(next)).toEqual(new Set([next.doc.child(0).attrs.paraId]));
    expect(hasStructuralChanges(next)).toBe(true);
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test("detects a count-neutral paragraph move", () => {
    const initial = trackedState(
      { text: "Moved", paraId: "11111111" },
      { text: "Kept", paraId: "22222222" },
    );
    const moved = initial.doc.child(0);
    const tr = initial.tr.delete(0, moved.nodeSize);
    tr.insert(tr.doc.content.size, moved);
    const next = initial.apply(tr);

    expect(next.doc.child(1).attrs.paraId).toBe("11111111");
    expect(hasStructuralChanges(next)).toBe(true);
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test("detects a table attribute change with stable paragraph membership", () => {
    const paragraph = schema.node("paragraph", { paraId: "11111111" }, schema.text("Cell"));
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [schema.node("tableCell", null, [paragraph])]),
    ]);
    const initial = EditorState.create({
      doc: schema.node("doc", null, [table]),
      plugins: [plugin, allocator],
    });
    const next = initial.apply(initial.tr.setNodeAttribute(0, "width", 5000));

    expect(hasStructuralChanges(next)).toBe(true);
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test("unrelated source ID backfill does not poison a new paragraph insertion", () => {
    const initial = trackedState({ text: "Unedited source" });
    const next = initial.apply(initial.tr.insert(0, schema.node("paragraph")));

    expect(getChangedParagraphIds(next)).toEqual(new Set([next.doc.child(0).attrs.paraId]));
    expect(hasUntrackedChanges(next)).toBe(false);
  });

  test("does not forgive editing an unidentified source paragraph after allocation", () => {
    const initial = trackedState({ text: "Unidentified source" });
    const next = initial.apply(initial.tr.insertText("!", 1));

    expect(next.doc.child(0).attrs.paraId).toMatch(/^[0-9A-F]{8}$/u);
    expect(hasUntrackedChanges(next)).toBe(true);
    const inserted = next.apply(next.tr.insert(0, schema.node("paragraph")));
    expect(hasUntrackedChanges(inserted)).toBe(true);
  });

  test("does not forgive deleting an unidentified source paragraph", () => {
    const initial = trackedState(
      { text: "Unidentified source" },
      { text: "Kept", paraId: "11111111" },
    );
    const next = initial.apply(initial.tr.delete(0, initial.doc.child(0).nodeSize));

    expect(hasUntrackedChanges(next)).toBe(true);
  });

  test("clear resets dirty paragraph ownership before subsequent allocation", () => {
    const initial = trackedState({ text: "First", paraId: "11111111" });
    const edited = initial.apply(initial.tr.insertText("!", 1));
    const cleared = edited.apply(clearTrackedChanges(edited));
    const next = cleared.apply(cleared.tr.insert(0, schema.node("paragraph")));

    expect(getChangedParagraphIds(next)).toEqual(new Set([next.doc.child(0).attrs.paraId]));
    expect(hasUntrackedChanges(next)).toBe(false);
  });
});
