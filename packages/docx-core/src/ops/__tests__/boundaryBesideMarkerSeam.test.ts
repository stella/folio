/**
 * A seam a removed insertion recorded, with only a zero-width marker across
 * it once the resolution is done, joins no pieces: a bookmark standing at it
 * stays where the document holds it, outside the content control after it.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import { OP_STORIES, REVISION_DECISIONS } from "../types";

const PARA_ID = "12345678";
const DATE = "2026-01-01T00:00:00Z";

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const paragraph = (content: ParagraphContent[]): Document => ({
  package: {
    document: { content: [{ type: "paragraph", paraId: PARA_ID, content }] },
  },
});

const contentOf = (document: Document): unknown => {
  const [block] = document.package.document.content;
  return block?.type === "paragraph" ? block.content : undefined;
};

const rejectAll = (document: Document, ids: readonly number[]): Document => {
  const resolved = applyDocumentOp(document, {
    type: "resolveRevision",
    story: OP_STORIES.MAIN,
    revisionIds: [...ids],
    decision: REVISION_DECISIONS.REJECT,
  });
  if (resolved.isErr()) throw resolved.error;
  return resolved.value.document;
};

/** Every revision id a value holds. */
const revisionIds = (value: unknown, out = new Set<number>()): number[] => {
  if (Array.isArray(value)) for (const item of value) revisionIds(item, out);
  else if (typeof value === "object" && value !== null)
    for (const [key, field] of Object.entries(value)) {
      if (
        key === "info" &&
        typeof field === "object" &&
        field !== null &&
        "id" in field &&
        typeof field.id === "number"
      )
        out.add(field.id);
      revisionIds(field, out);
    }
  return [...out];
};

test("a bookmark beside a seam with only a marker across it stays outside the control after it", () => {
  const original = paragraph([
    {
      type: "moveFromRangeStart",
      id: 1,
      name: "move1",
      author: "B",
      date: DATE,
    },
    { type: "moveFromRangeEnd", id: 1 },
    {
      type: "insertion",
      info: { id: 101, author: "B", date: DATE },
      content: [text("pq")],
    },
    { type: "bookmarkStart", id: 2, name: "mark2" },
    { type: "bookmarkEnd", id: 2 },
    {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 102 },
      content: [text("gh")],
    },
  ]);
  // A suggested link over "p" strikes it inside the insertion and inserts a
  // linked copy after it, cutting the insertion in two.
  const intent = {
    type: "setHyperlink" as const,
    from: { story: OP_STORIES.MAIN, blockId: PARA_ID, offset: 0 },
    to: { story: OP_STORIES.MAIN, blockId: PARA_ID, offset: 1 },
    href: "https://n.example/",
  };
  const ids = allocateEditorIntentIds(original, intent);
  const compiled = compileEditorIntent(original, {
    intent,
    mode: {
      type: "suggesting",
      revision: { id: ids.revisionId, author: "A", date: DATE },
      newIds: ids.newIds,
    },
  }).unwrap();
  const suggested = applyDocumentOps(original, compiled.ops).unwrap().document;
  // Rejecting every change leaves what rejecting the insertion leaves.
  expect(
    contentOf(rejectAll(suggested, revisionIds(contentOf(suggested)))),
  ).toEqual(contentOf(rejectAll(original, [101])));
  expect(contentOf(rejectAll(original, [101]))).toEqual([
    {
      type: "moveFromRangeStart",
      id: 1,
      name: "move1",
      author: "B",
      date: DATE,
    },
    { type: "moveFromRangeEnd", id: 1 },
    { type: "bookmarkStart", id: 2, name: "mark2" },
    { type: "bookmarkEnd", id: 2 },
    {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 102 },
      content: [text("gh")],
    },
  ]);
});
