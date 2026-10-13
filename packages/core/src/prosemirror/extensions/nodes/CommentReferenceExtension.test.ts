/**
 * The editor keeps one comment reference per commented range.
 *
 * The reference is the comment's visible mark, so an edit that leaves two of
 * them paints the mark twice and one that leaves none hides the comment. The
 * repair runs on every transaction rather than at each command, because the
 * commands that can break it are open-ended: a paste, a delete, a mark added
 * by an AI edit.
 */

import { describe, expect, test } from "bun:test";
import { Window } from "happy-dom";
import { EditorState, type Plugin } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";

import { schema } from "../../schema";
import { readCommentReferenceAttrs } from "../../commentReferenceAttrs";
import { planCommentReferenceRepairs } from "../../commentReferenceIntegrity";
import { TRACKED_RUN_INLINE_ATOM_DISPOSITIONS } from "../../trackedRunInlineAtoms";
import { CommentReferenceExtension } from "./CommentReferenceExtension";

const integrityPlugin = (): Plugin => {
  const plugin = CommentReferenceExtension().onSchemaReady({ schema }).plugins?.at(0);
  if (!plugin) {
    throw new Error("CommentReferenceExtension must enforce reference integrity");
  }
  return plugin;
};

const commentMark = (commentId: number) => schema.mark("comment", { commentId });
const reference = (commentId: number) => schema.node("commentReference", { commentId });

const referenceIds = (doc: PMNode): number[] => {
  const ids: number[] = [];
  doc.descendants((node) => {
    if (node.type.name === "commentReference") {
      ids.push(node.attrs["commentId"]);
    }
    return true;
  });
  return ids;
};

const stateWith = (paragraph: readonly PMNode[]): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [schema.node("paragraph", null, [...paragraph])]),
    plugins: [integrityPlugin()],
  });

describe("CommentReferenceExtension editing integrity", () => {
  test("keeps signed comment identities distinct from malformed references", () => {
    expect(readCommentReferenceAttrs(reference(-7))).toEqual({
      ok: true,
      value: { commentId: -7 },
    });
    expect(
      planCommentReferenceRepairs({
        references: [
          { status: "malformed", position: 9, size: 1 },
          { status: "valid", commentId: -7, position: 6, size: 1 },
        ],
        markExtents: [{ commentId: -7, end: 6 }],
      }),
    ).toEqual([{ type: "delete", position: 9, size: 1 }]);
  });

  test("reads a signed comment identity from editor DOM", () => {
    const getAttrs = schema.nodes.commentReference?.spec.parseDOM?.at(0)?.getAttrs;
    if (!getAttrs) {
      throw new Error("CommentReferenceExtension must read its editor DOM identity");
    }
    const element = new Window().document.createElement("span");
    element.setAttribute("data-docx-comment-reference", "-7");

    expect(getAttrs(element)).toEqual({ commentId: -7 });
  });

  test("gives a comment mark with no reference one after its last marked node", () => {
    const state = stateWith([schema.text("alpha")]);
    const applied = state.applyTransaction(
      state.tr.addMark(1, 6, commentMark(7)).addStoredMark(commentMark(7)),
    );

    expect(referenceIds(applied.state.doc)).toEqual([7]);
    const paragraph = applied.state.doc.firstChild;
    expect(paragraph?.lastChild?.type.name).toBe("commentReference");
  });

  test("drops a duplicate reference a paste left behind", () => {
    const state = stateWith([schema.text("alpha", [commentMark(3)]), reference(3)]);
    const applied = state.applyTransaction(state.tr.insert(6, reference(3)));

    expect(referenceIds(applied.state.doc)).toEqual([3]);
  });

  test("drops a reference whose comment no longer marks anything", () => {
    const state = stateWith([schema.text("alpha", [commentMark(4)]), reference(4)]);
    const applied = state.applyTransaction(state.tr.delete(1, 6));

    expect(referenceIds(applied.state.doc)).toEqual([]);
  });

  test("settles after one repair", () => {
    const state = stateWith([schema.text("alpha")]);
    const applied = state.applyTransaction(state.tr.addMark(1, 6, commentMark(9)));
    const again = applied.state.applyTransaction(applied.state.tr.insertText("!", 6));

    expect(referenceIds(again.state.doc)).toEqual([9]);
  });
});

type TransparentAtom = {
  [
    Name in keyof typeof TRACKED_RUN_INLINE_ATOM_DISPOSITIONS
  ]: (typeof TRACKED_RUN_INLINE_ATOM_DISPOSITIONS)[Name] extends "transparent" ? Name : never;
}[keyof typeof TRACKED_RUN_INLINE_ATOM_DISPOSITIONS];
const transparentAtoms = {
  commentReference: () => reference(7),
  rangeAnchor: () =>
    schema.node("rangeAnchor", {
      start: { type: "commentRangeStart", id: 7 },
      end: { type: "commentRangeEnd", id: 7 },
    }),
  renderedPageBreak: () => schema.node("renderedPageBreak"),
} satisfies Record<TransparentAtom, () => PMNode>;

// Comment ranges over marker atoms were not part of the editing-integrity fixture.
for (const [name, create] of Object.entries(transparentAtoms)) {
  test(`transparent ${name} cannot keep another comment range alive`, () => {
    const state = stateWith([
      schema.text("alpha", [commentMark(7)]),
      create(),
      schema.text("tail"),
    ]);
    const edited = state.applyTransaction(state.tr.addMark(1, 7, commentMark(23))).state;
    edited.doc.descendants((node) => {
      if (node.type.name === name)
        expect(node.marks.some(({ type }) => type.name === "comment")).toBe(false);
    });
    const deleted = edited.applyTransaction(edited.tr.delete(1, 6)).state;
    expect(referenceIds(deleted.doc)).not.toContain(23);
    const fixed = deleted.applyTransaction(
      deleted.tr.insertText("!", deleted.doc.content.size - 1),
    );
    expect(fixed.transactions).toHaveLength(1);
  });
}
