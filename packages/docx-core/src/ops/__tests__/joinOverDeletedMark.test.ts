/**
 * A suggested join of a paragraph whose mark a pending deletion already took:
 * the two already show as one paragraph, so the suggestion leaves that
 * deletion as it is and records only the properties the join gives.
 */
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { compileEditorIntent } from "../editorIntent";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

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

test.each(["Earlier Reviewer", "Reviewer"])(
  "a suggested join over a mark %s's pending deletion took keeps that deletion",
  (author) => {
    const deleted: Paragraph = {
      ...paragraph("00000001", "First.", "center"),
      pPrMark: {
        kind: "del",
        info: { id: 1, author, date: "2026-01-01T00:00:00Z" },
      },
    };
    const document: Document = {
      package: {
        document: {
          content: [deleted, paragraph("00000002", "Second.", "right")],
        },
      },
    };
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "joinParagraphs",
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
      },
      mode: {
        type: "suggesting",
        revision: {
          id: 100,
          author: "Reviewer",
          date: "2026-02-03T04:05:06Z",
        },
        newIds: { revision: [101, 102, 103] },
      },
    });
    if (compiled.isErr()) throw compiled.error;
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) throw applied.error;
    const [first] = storyParagraphs(applied.value.document.package.document);
    expect(first?.paragraph.pPrMark).toEqual(deleted.pPrMark);

    // Rejecting what the suggestion recorded gives the document back.
    const rejected = applyDocumentOps(applied.value.document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: applied.value.revisions,
        decision: REVISION_DECISIONS.REJECT,
      },
    ]);
    if (rejected.isErr()) throw rejected.error;
    expect(rejected.value.document).toStrictEqual(document);
  },
);
