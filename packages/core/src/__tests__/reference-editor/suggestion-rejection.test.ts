import { expect, mock, test } from "bun:test";
import { EditorState, type Transaction } from "prosemirror-state";

import { schema } from "../../prosemirror/schema";
import {
  rejectAIEditRevision,
  rejectAllSuggestions,
  rejectSuggestion,
} from "./comments";

const DATE = "2026-09-08T00:00:00.000Z";
const suggestedRevision = (revisionId: number, suggestionId: string) => ({
  revisionId,
  author: "Reviewer",
  date: DATE,
  provenance: "suggested",
  suggestionId,
});

const dispatchSpy = () => mock((_transaction: Transaction) => {});

test("targeted reference rejection refuses a revision shared by two suggestions", () => {
  const first = schema.mark("insertion", suggestedRevision(7, "first"));
  const other = schema.mark("insertion", suggestedRevision(7, "other"));
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("First ", [first]),
        schema.text("Other ", [other]),
        schema.text("Last", [first]),
      ]),
    ]),
  });
  const dispatch = dispatchSpy();

  expect(rejectSuggestion("first")(state, dispatch)).toBe(false);
  expect(dispatch).not.toHaveBeenCalled();
  expect(rejectAllSuggestions()(state, dispatch)).toBe(true);
  expect(dispatch).toHaveBeenCalledTimes(1);
  expect(dispatch.mock.calls.at(0)?.at(0)?.doc.textContent).toBe("");
});

test.each([
  ["targeted", rejectSuggestion("first")],
  ["bulk", rejectAllSuggestions()],
] as const)(
  "%s reference rejection leaves proposed comments intact when a join refuses",
  (mode, reject) => {
    const proposedComment = schema.mark("comment", {
      commentId: 23,
      suggestionId: "first",
    });
    const { revisionId: firstId, ...first } = suggestedRevision(7, "first");
    const other = suggestedRevision(8, "other");
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node(
          "paragraph",
          { pPrMark: { kind: "ins", info: { ...first, id: firstId } } },
          schema.text("Existing text", [proposedComment]),
        ),
        schema.node(
          "paragraph",
          {
            _suggestedInsert: other,
            pPrMark: {
              kind: "del",
              info: { id: 11, author: "Original editor", date: DATE },
            },
          },
          schema.text("Other proposal", [schema.mark("insertion", other)]),
        ),
      ]),
    });
    const dispatch = dispatchSpy();
    const ids = mode === "targeted" ? [firstId] : [firstId, other.revisionId];

    expect(rejectAIEditRevision(ids)(state, dispatch)).toBe(false);
    expect(reject(state, dispatch)).toBe(false);
    expect(dispatch).not.toHaveBeenCalled();
  },
);

test("a proposed comment rejects without an inline or structural revision", () => {
  const mark = schema.mark("comment", {
    commentId: 23,
    suggestionId: "comment",
  });
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, schema.text("Existing text", [mark])),
    ]),
  });
  const dispatch = dispatchSpy();

  expect(rejectSuggestion("comment")(state, dispatch)).toBe(true);
  const doc = dispatch.mock.calls.at(0)?.at(0)?.doc;
  expect(doc?.textContent).toBe("Existing text");
  expect(doc?.firstChild?.firstChild?.marks).toEqual([]);
});
