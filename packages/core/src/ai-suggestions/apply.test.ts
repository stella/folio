import { expect, test } from "bun:test";
import { EditorState, type Transaction } from "prosemirror-state";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import { mintRevisionId } from "../prosemirror/plugins/revisionIds";
import { schema } from "../prosemirror/schema";
import { applySuggestions } from "./apply";
import type { AISuggestion } from "./types";

test("suggestion batches share bounded revision allocation with editor commands", () => {
  const insertion = schema.marks["insertion"]!;
  const existingId = MAX_REVISION_ID;
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("old", [
        insertion.create({ revisionId: existingId, author: "Author" }),
      ]),
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

test("overlapping suggestions never change text outside the range each one covers", () => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text("abcdef")]),
  ]);
  const view = {
    state: EditorState.create({ doc }),
    dispatch: (tr: Transaction) => {
      view.state = view.state.apply(tr);
    },
  };
  const base = {
    topic: "Text",
    severity: "style",
    status: "pending",
    rationale: "Edit",
  } as const;
  const replaceBc = {
    ...base,
    id: "replace-bc",
    range: { from: 2, to: 4 },
    originalText: "bc",
    suggestedText: "BC",
    contextBefore: "a",
    contextAfter: "def",
  } satisfies AISuggestion;
  const deleteCd = {
    ...base,
    id: "delete-cd",
    range: { from: 3, to: 5 },
    originalText: "cd",
    suggestedText: "",
    contextBefore: "ab",
    contextAfter: "ef",
  } satisfies AISuggestion;

  const result = applySuggestions({
    view,
    suggestions: [replaceBc, deleteCd],
    mode: "direct",
    author: "Reviewer",
  });

  expect(result).toEqual({ applied: ["delete-cd"], stale: ["replace-bc"] });
  expect(view.state.doc.textContent).toBe("abef");
});

test("an insertion and a replacement at one start both apply, in either order", () => {
  const base = {
    topic: "Text",
    severity: "style",
    status: "pending",
    rationale: "Edit",
  } as const;
  const insert = {
    ...base,
    id: "insert",
    range: { from: 3, to: 3 },
    originalText: "",
    suggestedText: "X",
    contextBefore: "ab",
    contextAfter: "cdef",
  } satisfies AISuggestion;
  const replace = {
    ...base,
    id: "replace",
    range: { from: 3, to: 5 },
    originalText: "cd",
    suggestedText: "CD",
    contextBefore: "ab",
    contextAfter: "ef",
  } satisfies AISuggestion;
  for (const suggestions of [
    [insert, replace],
    [replace, insert],
  ]) {
    const view = {
      state: EditorState.create({
        doc: schema.node("doc", null, [
          schema.node("paragraph", null, [schema.text("abcdef")]),
        ]),
      }),
      dispatch: (tr: Transaction) => {
        view.state = view.state.apply(tr);
      },
    };
    const result = applySuggestions({
      view,
      suggestions,
      mode: "direct",
      author: "Reviewer",
    });
    expect(result.stale).toEqual([]);
    expect(view.state.doc.textContent).toBe("abXCDef");
  }
});
