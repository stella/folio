/**
 * Edits that join paragraphs report each removed paragraph with the one its
 * content went into: a join keeps the second paragraph's id.
 */
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { combinedAbsorptions } from "../edits";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const deleted = (paraId: string, text: string, id: number): Paragraph => ({
  ...paragraph(paraId, text),
  pPrMark: { kind: "del", info: { id, author: "A" } },
});
const documentOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});
const apply = (document: Document, op: DocumentOp) => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};
const resolve = (revisionIds: number[]): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
  story: OP_STORIES.MAIN,
  revisionIds,
  decision: REVISION_DECISIONS.ACCEPT,
});

test("accepting a deleted mark reports its paragraph joined into the next one", () => {
  const applied = apply(
    documentOf([deleted("00000001", "One ", 1), paragraph("00000002", "Two")]),
    resolve([1]),
  );
  expect(applied.touched.removed).toEqual(["00000001"]);
  expect(applied.touched.absorbed).toEqual([
    { absorbed: "00000001", survivor: "00000002" },
  ]);
});

test("a chain of deleted marks reports every removed paragraph against the paragraph left", () => {
  const applied = apply(
    documentOf([
      deleted("00000001", "One ", 1),
      deleted("00000002", "Two ", 2),
      paragraph("00000003", "Three"),
    ]),
    resolve([1, 2]),
  );
  expect(
    applied.touched.absorbed?.toSorted((left, right) =>
      left.absorbed.localeCompare(right.absorbed),
    ),
  ).toEqual([
    { absorbed: "00000001", survivor: "00000003" },
    { absorbed: "00000002", survivor: "00000003" },
  ]);
});

test("a direct join reports the paragraph it removed", () => {
  const applied = apply(
    documentOf([paragraph("00000001", "One "), paragraph("00000002", "Two")]),
    {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
    },
  );
  const [gone] = applied.touched.removed;
  expect(gone).toBeDefined();
  expect(applied.touched.absorbed).toEqual([
    {
      absorbed: gone ?? "",
      survivor: gone === "00000001" ? "00000002" : "00000001",
    },
  ]);
});

test("deleting one's own inserted paragraph reports its join into the next one", () => {
  const own: Paragraph = {
    ...paragraph("00000001", "Added"),
    pPrMark: { kind: "ins", info: { id: 9, author: "Reviewer" } },
  };
  const applied = apply(documentOf([own, paragraph("00000002", "Kept")]), {
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: ["00000001"],
    revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
    newIds: { revision: [101, 102, 103] },
  });
  expect(applied.touched.absorbed).toEqual([
    { absorbed: "00000001", survivor: "00000002" },
  ]);
});

test("absorptions made in turn chain to the paragraph holding them now", () => {
  expect(
    combinedAbsorptions(
      [
        [{ absorbed: "A", survivor: "B" }],
        [{ absorbed: "B", survivor: "C" }],
        [{ absorbed: "D", survivor: "E" }],
      ],
      new Set(["A", "B", "D", "E"]),
    ),
  ).toEqual([
    { absorbed: "A", survivor: "C" },
    { absorbed: "B", survivor: "C" },
  ]);
});
