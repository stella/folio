/**
 * Deleting a paragraph whose mark ends a section, right before a table: the
 * paragraph goes with its break, so its section and the one after it become
 * one under the later section's properties. Tracked, its words and its mark
 * are recorded deleted with the break on the mark; accepting matches the
 * direct deletion, rejecting gives the paragraph and its section back, and
 * every step has an exact inverse.
 */
import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  SectionProperties,
  Table,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  type DeleteBlocksOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const stamp = (id: number) => ({
  id,
  author: "Reviewer",
  date: "2026-02-03T04:05:06Z",
});

const ENDING: SectionProperties = {
  sectionStart: "oddPage",
  pageWidth: 11_000,
};
const FOLLOWING: SectionProperties = {
  sectionStart: "continuous",
  pageWidth: 12_000,
};

const paragraph = (
  paraId: string,
  text: string,
  sectionProperties?: SectionProperties,
): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(sectionProperties === undefined ? {} : { sectionProperties }),
});

const table: Table = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [paragraph("10000001", "Cell")] }],
    },
  ],
};

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const deletion = (tracked: boolean): DeleteBlocksOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  story: OP_STORIES.MAIN,
  blockIds: ["00000001"],
  ...(tracked && {
    revision: stamp(100),
    newIds: { revision: Array.from({ length: 16 }, (_, index) => 101 + index) },
  }),
});

const apply = (document: Document, op: DeleteBlocksOp) => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};

const undo = (
  document: Document,
  inverse: Parameters<typeof applyDocumentOps>[1],
) => {
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

describe.each([
  { place: "first in the body", before: [] },
  { place: "after a paragraph", before: [paragraph("00000009", "Lead")] },
])("a section-ending paragraph before a table, $place", ({ before }) => {
  const document = documentOf([
    ...before,
    paragraph("00000001", "Remove", ENDING),
    table,
    paragraph("00000002", "Tail", FOLLOWING),
  ]);
  const merged = documentOf([
    ...before,
    table,
    paragraph("00000002", "Tail", FOLLOWING),
  ]);

  test("goes directly with its break", () => {
    const direct = apply(document, deletion(false));
    expect(direct.document).toStrictEqual(merged);
    expect(undo(direct.document, direct.inverse)).toStrictEqual(document);
  });

  test("is tracked with the break on its deleted mark", () => {
    const tracked = apply(document, deletion(true));
    const removed = tracked.document.package.document.content.find(
      (block): block is Paragraph =>
        block.type === "paragraph" && block.paraId === "00000001",
    );
    expect(removed?.pPrMark?.kind).toBe("del");
    expect(removed?.sectionProperties).toStrictEqual(ENDING);
    expect(undo(tracked.document, tracked.inverse)).toStrictEqual(document);

    const accepted = resolve(
      tracked.document,
      tracked.revisions,
      REVISION_DECISIONS.ACCEPT,
    );
    expect(accepted.document).toStrictEqual(merged);
    expect(undo(accepted.document, accepted.inverse)).toStrictEqual(
      tracked.document,
    );

    const rejected = resolve(
      tracked.document,
      tracked.revisions,
      REVISION_DECISIONS.REJECT,
    );
    expect(rejected.document).toStrictEqual(document);
    expect(undo(rejected.document, rejected.inverse)).toStrictEqual(
      tracked.document,
    );
  });
});

test("alone before a story-final table, its mark still stays", () => {
  const document = documentOf([paragraph("00000001", "Remove", ENDING), table]);
  expect(applyDocumentOp(document, deletion(false)).isErr()).toBe(true);
});
