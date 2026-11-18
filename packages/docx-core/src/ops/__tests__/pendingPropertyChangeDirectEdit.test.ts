/**
 * A direct property edit of a paragraph whose tracked property change is
 * pending leaves that change recording what the paragraph now states, as a
 * reopened file reads it from the paragraph's own properties, and undoing
 * the edit restores the change exactly.
 */

import { describe, expect, mock, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import * as snapshots from "../paragraphPropertySnapshot";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

const documentOf = (...content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const firstParagraph = (document: Document): Paragraph | undefined =>
  document.package.document.content.find(
    (block): block is Paragraph => block.type === "paragraph",
  );

const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

/** A tracked alignment change, then a direct spacing edit; and its undo. */
const edits = () => {
  const source = documentOf({
    type: "paragraph",
    paraId: "00000001",
    formatting: { indentLeft: 720 },
    content: [],
  });
  const tracked = applied(source, {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    patch: { alignment: "center" },
    revision: { id: 1, author: "Reviewer" },
  }).document;
  const direct = applied(tracked, {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    patch: { spaceBefore: 240 },
  });
  let undone = direct.document;
  for (const inverse of direct.inverse)
    undone = applied(undone, inverse).document;
  return { tracked, direct: direct.document, undone };
};

/** Whether the direct edit left the pending change reading as reopened. */
const readsAsReopened = (): boolean => {
  const paragraph = firstParagraph(edits().direct);
  return (
    JSON.stringify(paragraph?.propertyChanges?.at(0)?.currentFormatting) ===
    JSON.stringify(paragraph?.formatting)
  );
};

describe("a direct edit after a tracked property change", () => {
  test("the pending change records what the paragraph now states", () => {
    const { direct } = edits();
    const paragraph = firstParagraph(direct);
    expect(paragraph?.formatting).toEqual({
      indentLeft: 720,
      alignment: "center",
      spaceBefore: 240,
    });
    expect(paragraph?.propertyChanges?.at(0)).toMatchObject({
      previousFormatting: { indentLeft: 720 },
      currentFormatting: paragraph?.formatting,
    });
    expect(readsAsReopened()).toBe(true);
  });

  test("undoing the direct edit restores the pending change exactly", () => {
    const { tracked, undone } = edits();
    expect(firstParagraph(undone)).toEqual(firstParagraph(tracked));
  });

  test("the check catches a direct edit that leaves the change as recorded", () => {
    // The test-only break: a direct edit keeps the pending change as the
    // tracked edit recorded it.
    const original = { ...snapshots };
    mock.module("../paragraphPropertySnapshot", () => ({
      ...original,
      changesRecordingAuthored: (
        changes: readonly NonNullable<Paragraph["propertyChanges"]>[number][],
      ) => [...changes],
    }));
    try {
      expect(readsAsReopened()).toBe(false);
    } finally {
      mock.module("../paragraphPropertySnapshot", () => original);
    }
  });
});
