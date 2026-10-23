import { GlobalRegistrator } from "@happy-dom/global-registrator";

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, test } from "bun:test";
import { DOMParser } from "prosemirror-model";
import {
  EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";

import { schema } from "../schema";
import {
  createSuggestionModePlugin,
  handleSuggestionPaste,
  suggestionModeKey,
} from "./suggestionMode";

const copiedWordHtml =
  '<p>keep <ins cite="mailto:Ann%20Author" datetime="2026-10-22T14:18">added</ins> ' +
  '<del cite="mailto:Bob%20Reviewer" datetime="2026-10-22T14:19">removed</del></p>';

test("tracked HTML paste keeps source deletions while proposing the rich slice", () => {
  const container = document.createElement("div");
  container.innerHTML = copiedWordHtml;
  const slice = DOMParser.fromSchema(schema).parseSlice(container);
  const original = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text("replace")]),
  ]);
  let state = EditorState.create({
    doc: original,
    selection: TextSelection.create(original, 1, 8),
    plugins: [createSuggestionModePlugin(true, "Reviewer")],
  });
  const view = {
    get state() {
      return state;
    },
    dispatch(transaction: Transaction) {
      state = state.apply(transaction);
    },
  };
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState) throw new Error("suggestion-mode state is missing");

  expect(handleSuggestionPaste(view, slice, pluginState)).toBe(true);
  const revisions: {
    text: string;
    kind: string;
    author: string;
    date: string | null;
    revisionId: number;
  }[] = [];
  state.doc.descendants((node) => {
    if (!node.isText) return;
    for (const mark of node.marks) {
      if (mark.type.name !== "insertion" && mark.type.name !== "deletion") {
        continue;
      }
      const author = mark.attrs["author"];
      const date = mark.attrs["date"];
      const revisionId = mark.attrs["revisionId"];
      if (
        typeof author !== "string" ||
        (typeof date !== "string" && date !== null) ||
        typeof revisionId !== "number"
      ) {
        continue;
      }
      revisions.push({
        text: node.text ?? "",
        kind: mark.type.name,
        author,
        date,
        revisionId,
      });
    }
  });

  const removed = revisions.find(
    ({ text, kind }) => text === "removed" && kind === "deletion",
  );
  expect(removed).toMatchObject({
    author: "Bob Reviewer",
    date: "2026-10-22T14:19",
  });
  expect(removed?.revisionId).toBeGreaterThan(0);
  expect(
    revisions.some(
      ({ text, kind, author }) =>
        text === "removed" && kind === "insertion" && author === "Reviewer",
    ),
  ).toBe(true);
  expect(
    revisions.some(
      ({ text, kind }) => text === "added" && kind === "insertion",
    ),
  ).toBe(true);
});

test("an untracked HTML paste remains parseable as an ordinary rich slice", () => {
  const container = document.createElement("div");
  container.innerHTML = "<p>plain <strong>rich</strong> text</p>";
  const slice = DOMParser.fromSchema(schema).parseSlice(container);
  expect(slice.content.textBetween(0, slice.content.size)).toBe(
    "plain rich text",
  );
});
