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
import { EditorState, type Plugin } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";

import { schema } from "../../schema";
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
