/**
 * A commit's projection is held as the document its transaction ends on, so
 * the next commit sees the state it was given is the projection without
 * comparing every node. A state that differs from the projection is still
 * refused as stale.
 */
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { schema } from "../prosemirror/schema";

const sessionWithState = () => {
  const document = fromMarkdown("First paragraph.\n\nSecond paragraph.");
  document.package.document.content.forEach((block, index) => {
    if (block.type === "paragraph")
      block.paraId = (index + 1).toString(16).padStart(8, "0");
  });
  const session = createCanonicalSession(document).unwrap();
  const state = EditorState.create({ schema, doc: session.projection.doc });
  return { session, state };
};

const typed = (
  session: ReturnType<typeof sessionWithState>["session"],
  state: EditorState,
) => {
  const at = 2;
  const selected = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at)),
  );
  const commit = session
    .prepareReplace(selected, { from: at, to: at, text: "x" })
    .unwrap();
  const next = selected.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  return { commit, next };
};

describe("a committed projection", () => {
  for (const mode of ["editing", "suggesting"] as const)
    test(`is held as the transaction's document, ${mode}`, () => {
      const { session, state } = sessionWithState();
      if (mode === "suggesting")
        session.setMode({ type: "suggesting", author: "Reviewer" });
      const { commit, next } = typed(session, state);
      expect(commit.projection.doc).toBe(next.doc);
      expect(session.projection.doc).toBe(next.doc);
      // A second commit from that state is accepted, and is held the same way.
      const again = typed(session, next);
      expect(session.projection.doc).toBe(again.next.doc);
    });

  test("refuses a state that differs from it", () => {
    const { session, state } = sessionWithState();
    const { next } = typed(session, state);
    const changed = next.apply(next.tr.insertText("y", 3));
    expect(
      session.prepareReplace(changed, { from: 2, to: 2, text: "z" }).isErr(),
    ).toBe(true);
  });
});
