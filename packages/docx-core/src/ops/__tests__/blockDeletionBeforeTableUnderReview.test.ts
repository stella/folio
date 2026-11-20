/**
 * Deleting a paragraph right before a table whose rows are still a pending
 * insertion: the deletion takes the paragraph's words and its mark, accepting
 * removes the emptied paragraph without crossing the table, and rejecting
 * gives it back. A paragraph the same author's pending insertion added whole
 * goes outright. A join that would run a paragraph into such a table, through
 * a paragraph whose mark already goes, is refused.
 */
import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import {
  type DeleteBlocksOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const stamp = (id: number, author: string) => ({
  id,
  author,
  date: "2026-02-03T04:05:06Z",
});

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** A table another author's (or `author`'s) pending insertion added. */
const insertedTable = (author: string): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      structuralChange: { type: "tableRowInsertion", info: stamp(60, author) },
      cells: [
        {
          type: "tableCell",
          content: [
            {
              ...paragraph("00000020", ""),
              pPrMark: { kind: "ins", info: stamp(61, author) },
              content: [
                {
                  type: "insertion",
                  info: stamp(62, author),
                  content: [
                    { type: "run", content: [{ type: "text", text: "Cell." }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  ],
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const deletion = (blockId: string, author = "Reviewer"): DeleteBlocksOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  story: OP_STORIES.MAIN,
  blockIds: [blockId],
  revision: stamp(100, author),
  newIds: { revision: Array.from({ length: 16 }, (_, index) => 101 + index) },
});

/** The leaves of a paragraph item: a run's, or the runs a revision holds. */
const leavesOf = (item: Paragraph["content"][number]) => {
  if (item.type === "run") return item.content;
  if (item.type !== "insertion" && item.type !== "deletion") return [];
  return item.content.flatMap((run) => (run.type === "run" ? run.content : []));
};

const texts = (document: Document) =>
  storyParagraphs(document.package.document).map(({ paragraph: placed }) =>
    placed.content
      .flatMap(leavesOf)
      .flatMap((leaf) => (leaf.type === "text" ? [leaf.text] : []))
      .join(""),
  );

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
  return resolved.value.document;
};

test.each(["Reviewer", "Other"])(
  "a paragraph before a table %s's pending insertion added loses its words and mark",
  (author) => {
    const document = documentOf([
      paragraph("00000001", "Intro."),
      paragraph("00000002", "Alpha."),
      insertedTable(author),
      paragraph("00000003", "Final."),
    ]);
    const deleted = applyDocumentOp(document, deletion("00000002"));
    if (deleted.isErr()) throw deleted.error;
    const removed = storyParagraphs(deleted.value.document.package.document)
      .map(({ paragraph: placed }) => placed)
      .find(({ paraId }) => paraId === "00000002");
    expect(removed?.pPrMark?.kind).toBe("del");

    const every = [60, 61, 62, ...deleted.value.revisions];
    expect(
      texts(resolve(deleted.value.document, every, REVISION_DECISIONS.ACCEPT)),
    ).toEqual(["Intro.", "Cell.", "Final."]);
    expect(
      texts(resolve(deleted.value.document, every, REVISION_DECISIONS.REJECT)),
    ).toEqual(["Intro.", "Alpha.", "Final."]);
    const undone = applyDocumentOps(
      deleted.value.document,
      deleted.value.inverse,
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(document);
  },
);

test("a paragraph its author added whole goes outright before that author's pending table", () => {
  const added: Paragraph = {
    ...paragraph("00000002", ""),
    pPrMark: { kind: "ins", info: stamp(50, "Reviewer") },
    content: [
      {
        type: "insertion",
        info: stamp(51, "Reviewer"),
        content: [{ type: "run", content: [{ type: "text", text: "Alpha." }] }],
      },
    ],
  };
  const document = documentOf([
    paragraph("00000001", "Intro."),
    added,
    insertedTable("Reviewer"),
    paragraph("00000003", "Final."),
  ]);
  const deleted = applyDocumentOp(document, deletion("00000002"));
  if (deleted.isErr()) throw deleted.error;
  expect(deleted.value.revisions).toEqual([]);
  expect(texts(deleted.value.document)).toEqual(["Intro.", "Cell.", "Final."]);
  const undone = applyDocumentOps(
    deleted.value.document,
    deleted.value.inverse,
  );
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
});

test.each([
  ["deleted", { kind: "del" as const, info: stamp(50, "Reviewer") }],
  [
    "inserted, then deleted",
    {
      kind: "ins" as const,
      info: stamp(50, "Other"),
      deletion: { info: stamp(51, "Reviewer") },
    },
  ],
])(
  "a tracked join through a paragraph whose mark is %s, before a table under review, is refused",
  (_shape, pPrMark) => {
    const document = documentOf([
      paragraph("00000001", "Intro."),
      { ...paragraph("00000002", "Alpha."), pPrMark },
      insertedTable("Reviewer"),
      paragraph("00000003", "Final."),
    ]);
    const joined = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      revision: stamp(100, "Reviewer"),
      newIds: { revision: [101, 102, 103] },
    });
    expect(joined.isErr() ? joined.error.reason : "applied").toBe(
      "untrackable",
    );
  },
);

test("an added paragraph holding a bookmark boundary is not removed outright", () => {
  const added: Paragraph = {
    ...paragraph("00000002", ""),
    pPrMark: { kind: "ins", info: stamp(50, "Reviewer") },
    content: [
      { type: "bookmarkStart", id: 5, name: "kept" },
      {
        type: "insertion",
        info: stamp(51, "Reviewer"),
        content: [{ type: "run", content: [{ type: "text", text: "Alpha." }] }],
      },
    ],
  };
  const document = documentOf([
    paragraph("00000001", "Intro."),
    added,
    insertedTable("Reviewer"),
    {
      ...paragraph("00000003", "Final."),
      content: [
        { type: "run", content: [{ type: "text", text: "Final." }] },
        { type: "bookmarkEnd", id: 5 },
      ],
    },
  ]);
  const deleted = applyDocumentOp(document, deletion("00000002"));
  if (deleted.isOk())
    expect(JSON.stringify(deleted.value.document)).toContain(
      '"type":"bookmarkStart"',
    );
});
