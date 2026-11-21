/**
 * Joining two paragraphs with range boundaries standing between them in the
 * block list: directly, the boundaries go into the joined paragraph where the
 * two meet; tracked, they stay between the blocks, and accepting the deleted
 * mark joins across them as the direct join does. Every step has an exact
 * inverse. A section endpoint or an inline merge depth still needs the two
 * to touch.
 */
import { describe, expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type JoinBlocksOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const join = (overrides: Partial<JoinBlocksOp> = {}): JoinBlocksOp => ({
  type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  nextBlockId: "00000002",
  ...overrides,
});

const STAMP = { id: 100, author: "Reviewer", date: "2026-01-01T00:00:00Z" };

const applied = (document: Document, op: JoinBlocksOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const undo = (document: Document, inverse: readonly DocumentOp[]) => {
  const undone = applyDocumentOps(document, inverse);
  if (undone.isErr()) throw undone.error;
  return undone.value.document;
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS],
) => {
  const resolved = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);
  if (resolved.isErr()) throw resolved.error;
  return resolved.value;
};

const SHAPES = {
  "a bookmark end": {
    before: [{ type: "bookmarkStart", id: 1, name: "Clause" }],
    standing: [{ type: "bookmarkEnd", id: 1 }],
    after: [],
  },
  "a bookmark start": {
    before: [],
    standing: [{ type: "bookmarkStart", id: 2, name: "Later" }],
    after: [{ type: "bookmarkEnd", id: 2 }],
  },
  "a bookmark end and a bookmark start": {
    before: [{ type: "bookmarkStart", id: 1, name: "Clause" }],
    standing: [
      { type: "bookmarkEnd", id: 1 },
      { type: "bookmarkStart", id: 2, name: "Later" },
    ],
    after: [{ type: "bookmarkEnd", id: 2 }],
  },
} as const satisfies Record<
  string,
  Record<"before" | "standing" | "after", readonly BlockContent[]>
>;

describe.each(Object.entries(SHAPES))(
  "a join across %s",
  (_, { before, standing, after }) => {
    const document = documentOf([
      ...before,
      paragraph("00000001", "Alpha"),
      ...standing,
      paragraph("00000002", "Beta"),
      ...after,
    ]);
    const joined = documentOf([
      ...before,
      {
        type: "paragraph",
        paraId: "00000002",
        content: [
          { type: "run", content: [{ type: "text", text: "Alpha" }] },
          ...standing,
          { type: "run", content: [{ type: "text", text: "Beta" }] },
        ],
      },
      ...after,
    ]);

    test("directly puts the boundaries where the paragraphs meet", () => {
      const direct = applied(document, join());
      expect(direct.document).toStrictEqual(joined);
      expect(undo(direct.document, direct.inverse)).toStrictEqual(document);
    });

    test("tracked keeps them between the blocks; accepting joins as directly", () => {
      const tracked = applied(document, join({ revision: STAMP }));
      expect(tracked.document.package.document.content).toHaveLength(
        document.package.document.content.length,
      );
      expect(undo(tracked.document, tracked.inverse)).toStrictEqual(document);
      for (const [decision, expected] of [
        [REVISION_DECISIONS.ACCEPT, joined],
        [REVISION_DECISIONS.REJECT, document],
      ] as const) {
        const resolved = resolve(tracked.document, tracked.revisions, decision);
        expect(resolved.document).toStrictEqual(expected);
        expect(undo(resolved.document, resolved.inverse)).toStrictEqual(
          tracked.document,
        );
      }
    });
  },
);

test("a section endpoint or an inline merge depth still needs the paragraphs to touch", () => {
  const document = documentOf([
    { ...paragraph("00000001", "Alpha"), sectionProperties: { pageWidth: 1 } },
    { type: "bookmarkStart", id: 2, name: "Later" },
    paragraph("00000002", "Beta"),
    { type: "bookmarkEnd", id: 2 },
  ]);
  expect(applyDocumentOp(document, join()).isErr()).toBe(true);
  const unsectioned = documentOf([
    paragraph("00000001", "Alpha"),
    { type: "bookmarkStart", id: 2, name: "Later" },
    paragraph("00000002", "Beta"),
    { type: "bookmarkEnd", id: 2 },
  ]);
  expect(applyDocumentOp(unsectioned, join({ depth: 1 })).isErr()).toBe(true);
});
