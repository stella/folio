import { describe, expect, test } from "bun:test";
import {
  EditorState,
  NodeSelection,
  type Transaction,
} from "prosemirror-state";

import { getCanonicalCommandIntents } from "../canonicalCommands";
import { schema } from "../schema";
import { replaceRangeText } from "./replaceRangeText";

/** The range of `word` in the one-paragraph document. */
const rangeOf = (state: EditorState, word: string) => {
  const from = 1 + state.doc.textContent.indexOf(word);
  return { from, to: from + word.length };
};

const legacy = (
  text: string,
  select: (state: EditorState) => { from: number; to: number },
) => {
  let state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Alpha Beta")]),
    ]),
  });
  const range = select(state);
  const dispatch = (transaction: Transaction) => {
    state = state.apply(transaction);
  };
  expect(replaceRangeText({ ...range, text })(state, dispatch)).toBe(true);
  return state;
};

describe("replaceRangeText", () => {
  test("deletes and replaces a range without a session", () => {
    expect(legacy("", (state) => rangeOf(state, " Beta")).doc.textContent).toBe(
      "Alpha",
    );
    expect(
      legacy("Gamma", (state) => rangeOf(state, "Beta")).doc.textContent,
    ).toBe("Alpha Gamma");
  });

  test("a collapsed empty replacement is no command", () => {
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("Alpha")]),
      ]),
    });
    const command = replaceRangeText({ from: 2, to: 2, text: "" });
    expect(command(state)).toBe(false);
    expect(getCanonicalCommandIntents(command, state)).toEqual([]);
  });

  test("a selected node is deleted as the selection", () => {
    let state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("A")]),
        schema.node("paragraph", null, [schema.text("B")]),
      ]),
    });
    state = state.apply(
      state.tr.setSelection(NodeSelection.create(state.doc, 3)),
    );
    const { from, to } = state.selection;
    expect(
      replaceRangeText({ from, to, text: "" })(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);
    expect(state.doc.textContent).toBe("A");
  });
});
