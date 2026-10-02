import { expect, test } from "bun:test";
import { MAX_REVISION_ID } from "@stll/docx-core/model";
import { EditorState, type Transaction } from "prosemirror-state";

import { schema } from "../prosemirror/schema";
import { mintRevisionId } from "../prosemirror/plugins/revisionIds";
import { applySuggestions } from "./apply";
import type { AISuggestion } from "./types";

test("suggestion batches share bounded revision allocation with editor commands", () => {
  const insertion = schema.marks["insertion"]!;
  const existingId = MAX_REVISION_ID;
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("old", [insertion.create({ revisionId: existingId, author: "Author" })]),
      schema.text(" target"),
    ]),
  ]);
  const view = {
    state: EditorState.create({ doc }),
    dispatch: (tr: Transaction) => {
      view.state = view.state.apply(tr);
    },
  };
  const suggestion = {
    id: "replace",
    topic: "Text",
    severity: "style",
    status: "pending",
    range: { from: 5, to: 11 },
    originalText: "target",
    suggestedText: "new",
    contextBefore: " ",
    contextAfter: "",
    rationale: "Replace text",
  } satisfies AISuggestion;
  const before = mintRevisionId();
  expect(
    applySuggestions({
      view,
      suggestions: [suggestion],
      mode: "tracked-changes",
      author: "Reviewer",
    }).applied,
  ).toEqual(["replace"]);
  const ids = new Set<number>();
  view.state.doc.descendants((node) => {
    for (const mark of node.marks) {
      const id: unknown = mark.attrs["revisionId"];
      if (typeof id === "number") ids.add(id);
    }
  });
  expect(ids.has(existingId)).toBe(true);
  ids.delete(existingId);
  expect(ids.size).toBeGreaterThan(0);
  for (const id of ids) {
    expect(Number.isInteger(id)).toBe(true);
    expect(id).toBeGreaterThan(0);
    expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
    expect(id).not.toBe(before);
  }
  expect(ids.has(mintRevisionId())).toBe(false);
});
