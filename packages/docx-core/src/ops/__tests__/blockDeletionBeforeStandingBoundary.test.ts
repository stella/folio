/**
 * Deleting a paragraph with range boundaries standing between it and the
 * next paragraph in the block list: the boundaries stand where the two
 * meet, as resolving the deleted mark joins them, so accepting the tracked
 * deletion is the direct one and rejecting it restores the original.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  type DeleteBlocksOp,
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
const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});
const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};
const deletion = (blockIds: string[]): DeleteBlocksOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  story: OP_STORIES.MAIN,
  blockIds,
  revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
  newIds: { revision: Array.from({ length: 32 }, (_, index) => 101 + index) },
});

/** Each block as its paragraph text pieces, or its boundary type. */
const shape = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content.map((item) =>
          item.type === "run"
            ? item.content
                .map((part) => (part.type === "text" ? part.text : ""))
                .join("")
            : item.type,
        )
      : block.type,
  );

const standings = {
  "a bookmark end": [{ type: "bookmarkEnd", id: 1 }],
  "a bookmark start": [{ type: "bookmarkStart", id: 2, name: "Later" }],
  "an end and a start": [
    { type: "bookmarkEnd", id: 1 },
    { type: "bookmarkStart", id: 2, name: "Later" },
  ],
} satisfies Record<string, BlockContent[]>;

for (const [name, standing] of Object.entries(standings))
  for (const deleted of [["00000002"], ["00000001", "00000002"]])
    test(`deleting ${deleted.length} paragraph(s) before ${name} joins the next paragraph across it`, () => {
      const opening = standing.some(({ type }) => type === "bookmarkEnd")
        ? [{ type: "bookmarkStart", id: 1, name: "Clause" } as const]
        : [];
      const closing = standing.some(({ type }) => type === "bookmarkStart")
        ? [{ type: "bookmarkEnd", id: 2 } as const]
        : [];
      const document = documentOf([
        ...opening,
        paragraph("00000001", "First"),
        paragraph("00000002", "Second"),
        ...standing,
        paragraph("00000003", "Third"),
        ...closing,
        paragraph("00000004", "Fourth"),
      ]);
      const op = deletion(deleted);
      const tracked = apply(document, op);
      const direct = apply(document, { ...op, revision: undefined });
      expect(shape(direct.document)).toEqual([
        ...opening.map(({ type }) => type),
        ...(deleted.length === 1 ? [["First"]] : []),
        [...standing.map(({ type }) => type), "Third"],
        ...closing.map(({ type }) => type),
        ["Fourth"],
      ]);
      for (const decision of Object.values(REVISION_DECISIONS)) {
        const resolved = apply(tracked.document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: tracked.revisions,
          decision,
        });
        expect(resolved.document).toStrictEqual(
          decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
        );
        const undo = applyDocumentOps(resolved.document, resolved.inverse);
        if (undo.isErr()) throw undo.error;
        expect(undo.value.document).toStrictEqual(tracked.document);
      }
      for (const edit of [tracked, direct]) {
        const undo = applyDocumentOps(edit.document, edit.inverse);
        if (undo.isErr()) throw undo.error;
        expect(undo.value.document).toStrictEqual(document);
      }
    });
