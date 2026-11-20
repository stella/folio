/**
 * A tracked join of a paragraph whose mark a pending deletion already took
 * keeps that deletion and records only the joined properties: accepting
 * leaves what the direct join leaves, rejecting leaves the original.
 */
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type JoinBlocksOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const paragraph = (
  paraId: string,
  text: string,
  alignment: "center" | "right",
): Paragraph => ({
  type: "paragraph",
  paraId,
  formatting: { alignment },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const earlier = {
  id: 1,
  author: "Earlier Reviewer",
  date: "2026-01-01T00:00:00Z",
};

test.each(["Earlier Reviewer", "Reviewer"])(
  "a tracked join over a mark %s's pending deletion took keeps that record",
  (author) => {
    const deleted: Paragraph = {
      ...paragraph("00000001", "First.", "center"),
      pPrMark: { kind: "del", info: { ...earlier, author } },
    };
    const document: Document = {
      package: {
        document: {
          content: [deleted, paragraph("00000002", "Second.", "right")],
        },
      },
    };
    const join: JoinBlocksOp = {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
      newIds: { revision: [101, 102, 103] },
    };
    const tracked = apply(document, join);
    const [first, second] = storyParagraphs(
      tracked.document.package.document,
    ).map(({ paragraph: placed }) => placed);
    expect(first?.pPrMark).toEqual(deleted.pPrMark);
    expect(second?.formatting).toEqual({ alignment: "center" });
    expect(second?.propertyChanges?.length).toBe(1);

    const direct = apply(document, { ...join, revision: undefined });
    const every = [earlier.id, ...tracked.revisions];
    const accepted = apply(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: every,
      decision: REVISION_DECISIONS.ACCEPT,
    });
    const shape = (edited: Document) =>
      storyParagraphs(edited.package.document).map(({ paragraph: placed }) => [
        placed.formatting,
        placed.content,
        placed.pPrMark,
        placed.propertyChanges,
      ]);
    expect(shape(accepted.document)).toEqual(shape(direct.document));
    const rejected = apply(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: every,
      decision: REVISION_DECISIONS.REJECT,
    });
    const rejectedOriginal = apply(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [earlier.id],
      decision: REVISION_DECISIONS.REJECT,
    });
    expect(shape(rejected.document)).toEqual(shape(rejectedOriginal.document));
    const undone = applyDocumentOps(tracked.document, tracked.inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(document);
  },
);
