import { describe, expect, test } from "bun:test";
import { history, redo, undo, undoDepth } from "prosemirror-history";
import { Schema } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Command, Transaction } from "prosemirror-state";

import {
  acceptChange,
  acceptAIEditRevision,
  findChangeAtPosition,
  findNextChange,
  findPreviousChange,
  rejectAIEditRevision,
  rejectChange,
} from "./comments";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      content: "inline*",
      group: "block",
      attrs: {
        numPr: { default: null },
        alignment: { default: null },
        _propertyChanges: { default: null },
        _sectionProperties: { default: null },
        pPrMark: { default: null },
      },
    },
    text: { group: "inline", marks: "_" },
    image: {
      group: "inline",
      inline: true,
      atom: true,
      marks: "_",
      toDOM: () => ["img"],
    },
  },
  marks: {
    // `bold` is here so a single tracked-change span can be split into
    // two adjacent text nodes that share the *same* insertion mark
    // instance but differ in inline formatting — ProseMirror does not
    // merge text nodes whose mark sets differ, so this is the only
    // way to construct a genuinely multi-node tracked change.
    bold: { toDOM: () => ["strong", 0] },
    insertion: {
      attrs: { revisionId: {}, author: {}, date: {} },
      excludes: "",
      toDOM: () => ["ins", 0],
    },
    deletion: {
      attrs: { revisionId: {}, author: {}, date: {} },
      excludes: "",
      toDOM: () => ["del", 0],
    },
  },
});

const REV_A_ATTRS = { revisionId: 1, author: "AI", date: "2026-01-01" };
const REV_B_ATTRS = { revisionId: 2, author: "AI", date: "2026-01-02" };

const makeStateWithMarks = () => {
  // One paragraph with two NON-overlapping insertion spans —
  // revision A's "alpha" then revision B's "beta", separated by a
  // plain word. Acting on revision A must leave "beta" and its
  // mark untouched; the previous resolveChange ignored the
  // revision id and processed every insertion mark in the range.
  const insertion = schema.marks["insertion"];
  return EditorState.create({
    schema,
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("alpha", [insertion.create(REV_A_ATTRS)]),
        schema.text(" middle "),
        schema.text("beta", [insertion.create(REV_B_ATTRS)]),
      ]),
    ]),
  });
};

const dispatcher = (state: EditorState) => {
  const view = {
    state,
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  return view;
};

const insertionRevisionsAt = (state: EditorState): number[] => {
  const ids: number[] = [];
  state.doc.descendants((node) => {
    if (!node.isText) {
      return;
    }
    for (const mark of node.marks) {
      if (mark.type.name === "insertion") {
        ids.push(mark.attrs["revisionId"] as number);
      }
    }
  });
  return ids;
};

describe("AI revision accept/reject scoping", () => {
  test("acceptAIEditRevision only touches marks for the matching revisionId", () => {
    const view = dispatcher(makeStateWithMarks());

    expect(insertionRevisionsAt(view.state)).toEqual([1, 2]);

    const ok = acceptAIEditRevision(REV_A_ATTRS.revisionId)(view.state, view.dispatch);

    expect(ok).toBe(true);
    // After accepting revision A, revision B's mark must still be
    // present — the user hasn't acted on it yet.
    expect(insertionRevisionsAt(view.state)).toEqual([2]);
  });

  test("rejectAIEditRevision only deletes text covered by the matching revisionId", () => {
    const view = dispatcher(makeStateWithMarks());

    const ok = rejectAIEditRevision(REV_A_ATTRS.revisionId)(view.state, view.dispatch);

    expect(ok).toBe(true);
    // Reject drops "alpha" (its inserted text) but must leave the
    // plain " middle " run and revision B's "beta" intact. The
    // remaining insertion mark belongs to revision B alone.
    expect(view.state.doc.textContent).toBe(" middle beta");
    expect(insertionRevisionsAt(view.state)).toEqual([2]);
  });
});

describe("findChangeAtPosition", () => {
  test("returns the input range unchanged for a non-cursor selection", () => {
    const state = makeStateWithMarks();
    expect(findChangeAtPosition(state, 3, 12)).toEqual({ from: 3, to: 12 });
  });

  test("returns the cursor as-is when no tracked-change marks exist nearby", () => {
    // Position 10 sits in the plain " middle " run between the two
    // insertion spans of `makeStateWithMarks`.
    const state = makeStateWithMarks();
    expect(findChangeAtPosition(state, 10, 10)).toEqual({
      from: 10,
      to: 10,
    });
  });

  test("expands a cursor inside a single tracked-change span to that span's full range", () => {
    // Cursor inside "alpha" (revision A) — expansion should cover the
    // whole inserted word. Paragraph offsets: 1..6 = "alpha".
    const state = makeStateWithMarks();
    expect(findChangeAtPosition(state, 3, 3)).toEqual({ from: 1, to: 6 });
  });

  test("expands across multiple adjacent text nodes that share the same mark instance", () => {
    // Three adjacent text nodes that all share the *same* insertion
    // mark — this comes up when the inserted span is formatted
    // unevenly (e.g., a bold word inside a tracked-change). The
    // expansion must reach both edges, not just one neighbouring
    // node on either side.
    const insertion = schema.marks["insertion"]!.create(REV_A_ATTRS);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("one", [insertion]),
          schema.text("two", [insertion]),
          schema.text("three", [insertion]),
        ]),
      ]),
    });
    // "one" 1..4, "two" 4..7, "three" 7..12. Cursor in "two".
    expect(findChangeAtPosition(state, 5, 5)).toEqual({ from: 1, to: 12 });
  });

  test("does not bleed into an adjacent insertion that belongs to a different revisionId", () => {
    // Two insertions back-to-back, no plain text between them. Cursor
    // is in the FIRST insertion. The expansion must stop at the
    // boundary — accepting/rejecting one revision must not implicate
    // the other.
    const insertion = schema.marks["insertion"];
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("alpha", [insertion.create(REV_A_ATTRS)]),
          schema.text("beta", [insertion.create(REV_B_ATTRS)]),
        ]),
      ]),
    });
    // "alpha" occupies positions 1..6, "beta" positions 6..10.
    expect(findChangeAtPosition(state, 3, 3)).toEqual({ from: 1, to: 6 });
  });

  test("expands a cursor inside a list property suggestion to the paragraph boundary", () => {
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node(
          "paragraph",
          {
            numPr: { numId: 1, ilvl: 0 },
            _propertyChanges: [
              {
                type: "paragraphPropertyChange",
                info: { id: 10, author: "Alice", date: "2026-01-01" },
                previousFormatting: { numPr: null },
              },
            ],
          },
          [schema.text("alpha")],
        ),
      ]),
    });

    expect(findChangeAtPosition(state, 3, 3)).toEqual({ from: 6, to: 7 });
  });

  test("accepts a property-only list suggestion from a cursor-derived range", () => {
    const initial = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node(
          "paragraph",
          {
            numPr: { numId: 1, ilvl: 0 },
            _propertyChanges: [
              {
                type: "paragraphPropertyChange",
                info: { id: 10, author: "Alice", date: "2026-01-01" },
                previousFormatting: { numPr: null },
              },
            ],
          },
          [schema.text("alpha")],
        ),
      ]),
    });
    const range = findChangeAtPosition(initial, 3, 3);
    const view = dispatcher(initial);

    expect(acceptChange(range.from, range.to)(view.state, view.dispatch)).toBe(true);

    expect(view.state.doc.child(0).attrs["_propertyChanges"]).toBeNull();
  });

  test("rejecting a pPrChange restores only in-scope properties and preserves out-of-scope attrs", () => {
    // Word stores a pPrChange's old properties as CT_PPrBase, which cannot
    // carry an inline sectPr (or its header/footer references). Rejecting a
    // paragraph-property change must therefore MERGE the stored old
    // properties over the live attrs, not replace the pPr wholesale — a
    // wholesale replace would drop the separately-modeled inline section
    // break. Regression guard for the OOXML corner case.
    const sectionProperties = {
      type: "nextPage",
      headerReferences: [{ type: "default", relationshipId: "rId7" }],
    };
    const initial = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node(
          "paragraph",
          {
            alignment: "left",
            _sectionProperties: sectionProperties,
            _propertyChanges: [
              {
                type: "paragraphPropertyChange",
                info: { id: 10, author: "Alice", date: "2026-01-01" },
                previousFormatting: { alignment: "center" },
              },
            ],
          },
          [schema.text("alpha")],
        ),
      ]),
    });
    const view = dispatcher(initial);

    expect(rejectChange(0, initial.doc.content.size)(view.state, view.dispatch)).toBe(true);

    const attrs = view.state.doc.child(0).attrs;
    expect(attrs["alignment"]).toBe("center");
    expect(attrs["_propertyChanges"]).toBeNull();
    expect(attrs["_sectionProperties"]).toEqual(sectionProperties);
  });
});

describe("tracked change navigation", () => {
  test("findNextChange returns the nearest later change before wrapping", () => {
    const state = makeStateWithMarks();

    const result = findNextChange(state, 6);

    expect(result).toMatchObject({
      from: 14,
      to: 18,
      type: "insertion",
    });
  });

  test("findNextChange wraps only when there is no later change", () => {
    const state = makeStateWithMarks();

    const result = findNextChange(state, state.doc.content.size);

    expect(result).toMatchObject({
      from: 1,
      to: 6,
      type: "insertion",
    });
  });

  test("findPreviousChange returns the nearest earlier change before wrapping", () => {
    const state = makeStateWithMarks();

    const result = findPreviousChange(state, 12);

    expect(result).toMatchObject({
      from: 1,
      to: 6,
      type: "insertion",
    });
  });

  test("findPreviousChange wraps only when there is no earlier change", () => {
    const state = makeStateWithMarks();

    const result = findPreviousChange(state, 0);

    expect(result).toMatchObject({
      from: 14,
      to: 18,
      type: "insertion",
    });
  });

  // A tracked-change span split into two text nodes that share the *same*
  // insertion mark instance but differ in inline formatting — exactly
  // what happens when a user partially bolds inside an AI suggestion.
  // The navigation buttons must select the WHOLE inserted span, not
  // just the first text node — otherwise pressing "next change" leaves
  // the second half outside the selection and follow-up accept/reject
  // actions miss it.
  const makeMultiNodeSameRevisionState = () => {
    const insertion = schema.marks["insertion"]!.create(REV_A_ATTRS);
    const bold = schema.marks["bold"]!.create();
    return EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("plain"),
          schema.text("bold", [insertion, bold]),
          schema.text("plain", [insertion]),
          schema.text(" tail"),
        ]),
      ]),
    });
  };

  test("findNextChange spans the full multi-node tracked change", () => {
    // Paragraph offsets: "plain" 1..6, "bold" 6..10 (ins+bold),
    // "plain" 10..15 (ins), " tail" 15..20. Starting before the
    // change, "next" must select the whole inserted region 6..15.
    const state = makeMultiNodeSameRevisionState();
    expect(findNextChange(state, 0)).toMatchObject({
      from: 6,
      to: 15,
      type: "insertion",
    });
  });

  test("findPreviousChange spans the full multi-node tracked change", () => {
    // Starting from beyond the change, "previous" must also select
    // the whole inserted region 6..15.
    const state = makeMultiNodeSameRevisionState();
    expect(findPreviousChange(state, 20)).toMatchObject({
      from: 6,
      to: 15,
      type: "insertion",
    });
  });

  test("findPreviousChange surfaces an overlapping change inside a previously expanded range", () => {
    // A previously inserted span that a second reviewer later deleted
    // ("inserted-then-deleted") carries BOTH an insertion mark
    // (revision A) and a deletion mark (revision B) on the same text
    // node. The forward walk must process both ranges — skipping
    // strictly by position would return the outer insertion when the
    // user expects the nearer overlapping deletion.
    const insertion = schema.marks["insertion"]!.create(REV_A_ATTRS);
    const deletion = schema.marks["deletion"]!.create(REV_B_ATTRS);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          // "first " — insertion A only, positions 1..7.
          schema.text("first ", [insertion]),
          // "second" — insertion A + deletion B, positions 7..13.
          schema.text("second", [insertion, deletion]),
        ]),
      ]),
    });
    expect(findPreviousChange(state, 14)).toMatchObject({
      from: 7,
      to: 13,
      type: "deletion",
    });
  });

  test("findNextChange returns the full range when startPos sits inside a multi-node change", () => {
    // The toolbar calls `findNextChange(state, selectionEnd)` and then
    // accepts/rejects the returned range. If `startPos` lands inside
    // an existing tracked change, returning a range clamped to
    // `startPos` truncates the earlier portion of the same revision
    // and leaves orphan marks after accept. Return the full expanded
    // range — including positions BEFORE `startPos`.
    const state = makeMultiNodeSameRevisionState();
    expect(findNextChange(state, 8)).toMatchObject({
      from: 6,
      to: 15,
      type: "insertion",
    });
  });

  test("findNextChange expands across marked inline atoms in the same revision", () => {
    const insertion = schema.marks["insertion"]!.create(REV_A_ATTRS);
    const image = schema.nodes["image"]!.create(null, null, [insertion]);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("text", [insertion]),
          image,
          schema.text("tail", [insertion]),
        ]),
      ]),
    });

    expect(findNextChange(state, 0)).toMatchObject({
      from: 1,
      to: 10,
      type: "insertion",
    });
  });
});

/**
 * Undoing / redoing an id-scoped accept/reject IN PLACE, through
 * prosemirror-history — the mechanism the editor ref's `undo()` / `redo()`
 * drive (ports upstream docx-editor ff971a7b). folio's id-scoped resolver is
 * `acceptAIEditRevision` / `rejectAIEditRevision`; each dispatches a single
 * transaction, so one `undo()` reverses the whole resolution — even a coalesced
 * revision whose sites span multiple paragraphs — with no document reload.
 */
describe("undo / redo of an id-scoped resolution — in-place reversal, no reload", () => {
  const REV = 7;

  /** One paragraph: `keep `, then `tracked` carrying `markName`@REV, then ` tail`. */
  const docWith = (markName: "insertion" | "deletion"): PMNode => {
    const markType = schema.marks[markName]!;
    const attrs = { revisionId: REV, author: "AI", date: "2026-01-01" };
    return schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("keep "),
        schema.text("tracked", [markType.create(attrs)]),
        schema.text(" tail"),
      ]),
    ]);
  };

  /** An editor with history(), so undo/redo run the ref's real mechanism. */
  const editor = (doc: PMNode) => {
    let state = EditorState.create({ schema, doc, plugins: [history()] });
    return {
      get state() {
        return state;
      },
      run(cmd: Command) {
        return cmd(state, (tr) => {
          state = state.apply(tr);
        });
      },
    };
  };

  const text = (state: EditorState) => state.doc.textContent;
  const hasMark = (state: EditorState, name: string) => {
    let found = false;
    state.doc.descendants((node) => {
      if (node.marks.some((mark) => mark.type.name === name)) {
        found = true;
      }
    });
    return found;
  };

  test("undo restores an accepted insertion mark; redo drops it again", () => {
    const ed = editor(docWith("insertion"));
    expect(ed.run(acceptAIEditRevision(REV))).toBe(true);
    expect(hasMark(ed.state, "insertion")).toBe(false);

    ed.run(undo);
    expect(hasMark(ed.state, "insertion")).toBe(true);
    expect(text(ed.state)).toBe("keep tracked tail");

    ed.run(redo);
    expect(hasMark(ed.state, "insertion")).toBe(false);
  });

  test("undo restores text removed by accepting a deletion", () => {
    const ed = editor(docWith("deletion"));
    expect(ed.run(acceptAIEditRevision(REV))).toBe(true);
    expect(text(ed.state)).toBe("keep  tail");

    ed.run(undo);
    expect(text(ed.state)).toBe("keep tracked tail");
    expect(hasMark(ed.state, "deletion")).toBe(true);
  });

  test("undo restores text + mark removed by rejecting an insertion; redo re-removes", () => {
    const ed = editor(docWith("insertion"));
    expect(ed.run(rejectAIEditRevision(REV))).toBe(true);
    expect(text(ed.state)).toBe("keep  tail");
    expect(hasMark(ed.state, "insertion")).toBe(false);

    ed.run(undo);
    expect(text(ed.state)).toBe("keep tracked tail");
    expect(hasMark(ed.state, "insertion")).toBe(true);

    ed.run(redo);
    expect(text(ed.state)).toBe("keep  tail");
  });

  test("resolves every site of a coalesced revision in one undoable step", () => {
    const insertion = schema.marks["insertion"]!;
    const attrs = { revisionId: REV, author: "AI", date: "2026-01-01" };
    // The same revisionId on two non-contiguous runs across two paragraphs.
    const ed = editor(
      schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("a "),
          schema.text("one", [insertion.create(attrs)]),
          schema.text(" b"),
        ]),
        schema.node("paragraph", null, [
          schema.text("c "),
          schema.text("two", [insertion.create(attrs)]),
          schema.text(" d"),
        ]),
      ]),
    );
    const before = undoDepth(ed.state);
    expect(ed.run(acceptAIEditRevision(REV))).toBe(true);
    expect(hasMark(ed.state, "insertion")).toBe(false); // every site resolved...
    expect(undoDepth(ed.state)).toBe(before + 1); // ...in a single history step,
    ed.run(undo);
    expect(hasMark(ed.state, "insertion")).toBe(true); // so one undo restores them all.
  });

  test("a selection-only change is never recorded, so undo() still targets the last edit", () => {
    const ed = editor(docWith("insertion"));
    ed.run(acceptAIEditRevision(REV));
    expect(hasMark(ed.state, "insertion")).toBe(false);
    // Move the selection — no document steps, so prosemirror-history ignores it
    // (the ref undo()'s documented guarantee that scroll/locate isn't undone).
    ed.run((state, dispatch) => {
      dispatch?.(state.tr.setSelection(TextSelection.create(state.doc, 1)));
      return true;
    });
    ed.run(undo);
    expect(hasMark(ed.state, "insertion")).toBe(true); // reverted the ACCEPT, not the selection
  });
});
