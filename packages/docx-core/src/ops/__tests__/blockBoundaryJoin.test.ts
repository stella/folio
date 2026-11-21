/**
 * Resolving away a paragraph mark with range boundaries standing between
 * the paragraph and the next one in the block list joins the two, the
 * boundaries where they meet: what stood before them stays before.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (
  paraId: string,
  text: string,
  mark?: "ins" | "del",
): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(mark === undefined
    ? {}
    : {
        pPrMark: {
          kind: mark,
          info: { id: 9, author: "A", date: "2026-01-02T03:04:05Z" },
        },
      }),
});

const documentOf = (mark: "ins" | "del"): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", "First", mark),
        { type: "bookmarkStart", id: 1, name: "Clause" },
        paragraph("00000002", "Second"),
        { type: "bookmarkEnd", id: 1 },
        paragraph("00000003", "Third"),
      ] satisfies BlockContent[],
    },
  },
});

const resolved = (document: Document, decision: RevisionDecision) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [9],
    decision,
  });
  if (applied.isErr()) throw applied.error;
  const undo = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(document);
  return applied.value.document.package.document.content;
};

for (const [mark, decision] of [
  ["del", REVISION_DECISIONS.ACCEPT],
  ["ins", REVISION_DECISIONS.REJECT],
] as const) {
  test(`${decision}ing a mark before a block-level bookmark start joins across it`, () => {
    const [joined, ...rest] = resolved(documentOf(mark), decision);
    expect(joined?.type).toBe("paragraph");
    expect(
      joined?.type === "paragraph"
        ? joined.content.map((item) =>
            item.type === "run"
              ? item.content
                  .map((part) => (part.type === "text" ? part.text : ""))
                  .join("")
              : item.type,
          )
        : [],
    ).toEqual(["First", "bookmarkStart", "Second"]);
    expect(rest.map((block) => block.type)).toEqual([
      "bookmarkEnd",
      "paragraph",
    ]);
  });
}

// Joining the paragraphs directly does the same, and undoing it stands the
// boundary between them again; tracked, the boundary stays until the join
// resolves, accepting as the direct join does and rejecting as nothing did.
test("joining paragraphs across a block-level bookmark start meets the laws", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph("00000001", "First"),
          { type: "bookmarkStart", id: 1, name: "Clause" },
          paragraph("00000002", "Second"),
          { type: "bookmarkEnd", id: 1 },
        ] satisfies BlockContent[],
      },
    },
  };
  const join = {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    nextBlockId: "00000002",
  } as const;
  const applied = (op: Parameters<typeof applyDocumentOp>[1]) => {
    const result = applyDocumentOp(document, op);
    if (result.isErr()) throw result.error;
    const undo = applyDocumentOps(result.value.document, result.value.inverse);
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(document);
    return result.value;
  };
  const direct = applied(join);
  expect(direct.document.package.document.content).toEqual([
    {
      ...paragraph("00000002", "Second"),
      content: [
        ...paragraph("", "First").content,
        { type: "bookmarkStart", id: 1, name: "Clause" },
        ...paragraph("", "Second").content,
      ],
    },
    { type: "bookmarkEnd", id: 1 },
  ]);
  const tracked = applied({
    ...join,
    revision: { id: 20, author: "B", date: "2026-01-03T00:00:00Z" },
  });
  expect(tracked.document.package.document.content.at(1)).toEqual({
    type: "bookmarkStart",
    id: 1,
    name: "Clause",
  });
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const resolution = applyDocumentOp(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [20],
      decision,
    });
    if (resolution.isErr()) throw resolution.error;
    expect(resolution.value.document).toStrictEqual(
      decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
    );
  }
});
