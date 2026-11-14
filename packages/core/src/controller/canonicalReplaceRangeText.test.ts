import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { replaceRangeText } from "../prosemirror/commands/replaceRangeText";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import { prepareCanonicalCommands } from "./canonicalStructure";

const source = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [
            { type: "run", content: [{ type: "text", text: "Alpha Beta" }] },
          ],
        },
      ],
    },
  },
});

/** The range of `word` in the one-paragraph document. */
const rangeOf = (state: EditorState, word: string) => {
  const from = 1 + state.doc.textContent.indexOf(word);
  return { from, to: from + word.length };
};

const canonicalText = (
  text: string,
  replace: (state: EditorState) => { from: number; to: number },
  suggesting = false,
) => {
  const session = createCanonicalSession(source()).unwrap();
  if (suggesting) session.setMode({ type: "suggesting", author: "Ada" });
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const range = replace(state);
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, range.from, range.to),
    ),
  );
  const command = replaceRangeText({ ...range, text });
  const intents = getCanonicalCommandIntents(command, state);
  expect(intents).toEqual([{ type: "replaceText", ...range, text }]);
  const commit = prepareCanonicalCommands(session, state, intents ?? []);
  if (commit.isErr()) throw commit.error;
  expect(commit.value.publish().isOk()).toBe(true);
  return session;
};

describe("replaceRangeText in a canonical session", () => {
  test("a canonical session journals the replacement as its own text input", () => {
    const deleted = canonicalText("", (state) => rangeOf(state, " Beta"));
    expect(deleted.projection.doc.textContent).toBe("Alpha");
    expect(deleted.canUndo).toBe(true);
    const replaced = canonicalText("Gamma", (state) => rangeOf(state, "Beta"));
    expect(replaced.projection.doc.textContent).toBe("Alpha Gamma");
  });

  test("a suggesting canonical session keeps the deleted text as a suggestion", () => {
    const session = canonicalText("", (state) => rangeOf(state, "Beta"), true);
    // The text stays on the page, marked for deletion.
    expect(session.projection.doc.textContent).toBe("Alpha Beta");
    let deletion = false;
    session.projection.doc.descendants((node) => {
      if (node.marks.some((mark) => mark.type.name === "deletion"))
        deletion = true;
    });
    expect(deletion).toBe(true);
  });

  test("the same edit without its descriptor has no intent for the session", () => {
    const session = createCanonicalSession(source()).unwrap();
    const state = EditorState.create({ schema, doc: session.projection.doc });
    const range = rangeOf(state, "Beta");
    const command = replaceRangeText({ ...range, text: "" });
    // The menus' former in-place edit: the same change, with no descriptor
    // for the view's command owner to journal, so a session refuses it.
    const undescribed: typeof command = (current, dispatch, view) =>
      command(current, dispatch, view);
    expect(getCanonicalCommandIntents(undescribed, state)).toBeUndefined();
    expect(getCanonicalCommandIntents(command, state)).toEqual([
      { type: "replaceText", ...range, text: "" },
    ]);
  });
});
