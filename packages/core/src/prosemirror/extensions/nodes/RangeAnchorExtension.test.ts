/**
 * How a point comment behaves under the caret.
 *
 * The anchor is a zero-width atom, so the two rules a reviewer would expect
 * both fall out of ProseMirror's own position mapping rather than out of a
 * command that has to remember them:
 *
 * - typing beside the anchor does not widen the range. A point comment stays a
 *   point unless the reviewer selects text and comments on the selection.
 * - deleting the anchor removes the comment, exactly as deleting the reference
 *   does today: the anchor is the only thing naming the comment, so the
 *   reference-integrity pass reads the reference as an orphan and drops it.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Plugin } from "prosemirror-state";

import { schema } from "../../schema";
import { CommentReferenceExtension } from "./CommentReferenceExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./RangeAnchorExtension";

const COMMENT_ID = 7;

const integrityPlugin = (): Plugin => {
  const plugin = CommentReferenceExtension().onSchemaReady({ schema }).plugins?.at(0);
  if (!plugin) {
    throw new Error("CommentReferenceExtension must enforce reference integrity");
  }
  return plugin;
};

const anchor = (): PMNode =>
  schema.node(RANGE_ANCHOR_NODE_NAME, {
    start: { type: "commentRangeStart", id: COMMENT_ID },
    end: { type: "commentRangeEnd", id: COMMENT_ID },
  });

const reference = (): PMNode => schema.node("commentReference", { commentId: COMMENT_ID });

const nodeNames = (doc: PMNode): string[] => {
  const names: string[] = [];
  doc.descendants((node) => {
    names.push(node.isText ? `text:${node.text ?? ""}` : node.type.name);
    return !node.isAtom;
  });
  return names;
};

/** `alpha` · anchor · reference · `omega` */
const stateWithAnchor = (): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("alpha"),
        anchor(),
        reference(),
        schema.text("omega"),
      ]),
    ]),
    plugins: [integrityPlugin()],
  });

const anchorPosition = (doc: PMNode): number => {
  let position = -1;
  doc.descendants((node, at) => {
    if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
      position = at;
    }
    return true;
  });
  return position;
};

describe("a point comment under the caret", () => {
  test("typing at the anchor does not widen the range", () => {
    const state = stateWithAnchor();
    const at = anchorPosition(state.doc);
    const applied = state.applyTransaction(state.tr.insertText("typed", at));

    // The typed text carries no comment mark, so the range is still the anchor
    // and still spans nothing.
    let marked = 0;
    applied.state.doc.descendants((node) => {
      if (node.marks.some((mark) => mark.type.name === "comment")) {
        marked += 1;
      }
      return true;
    });
    expect(marked).toBe(0);
    expect(nodeNames(applied.state.doc)).toEqual([
      "paragraph",
      "text:alphatyped",
      RANGE_ANCHOR_NODE_NAME,
      "commentReference",
      "text:omega",
    ]);
  });

  test("deleting the anchor removes the comment", () => {
    const state = stateWithAnchor();
    const at = anchorPosition(state.doc);
    const applied = state.applyTransaction(state.tr.delete(at, at + 1));

    expect(nodeNames(applied.state.doc)).toEqual(["paragraph", "text:alphaomega"]);
  });

  test("the anchor alone keeps the reference the integrity pass would drop", () => {
    const state = stateWithAnchor();
    const applied = state.applyTransaction(state.tr.insertText("!", 1));

    expect(nodeNames(applied.state.doc)).toContain("commentReference");
    expect(nodeNames(applied.state.doc)).toContain(RANGE_ANCHOR_NODE_NAME);
  });
});
