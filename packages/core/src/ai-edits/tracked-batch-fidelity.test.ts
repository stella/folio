import { expect, test } from "bun:test";

import {
  paragraphsDocx,
  reviewComment,
  textRun,
} from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

const commentedBlockId = "10000000";
const secondBlockId = "10000002";
const commentedText = "Document heading";
const originalText = "Original wording";
const replacementText = "Revised wording";

const source = () =>
  paragraphsDocx(
    [
      [
        { type: "commentRangeStart", id: 1 },
        textRun(commentedText, { bold: true }),
        { type: "commentRangeEnd", id: 1 },
        { type: "commentReference", id: 1 },
      ],
      [textRun("Intervening paragraph.")],
      [textRun(originalText)],
    ],
    [reviewComment(1, "Check heading.")],
  );

const edit = {
  id: "replace-wording",
  type: "replaceInBlock",
  blockId: secondBlockId,
  find: originalText,
  replace: replacementText,
} as const;

for (const session of ["legacy", "canonical"] as const) {
  test(`${session} refuses the entire tracked batch when accepting would remove a comment`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
      session,
    });
    const original = reviewer.toDocument();
    const comments = reviewer.getComments();
    expect(comments.map(({ anchoredText }) => anchoredText)).toEqual([
      commentedText,
    ]);
    const batch = {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      atomic: false,
      operations: [
        {
          id: "delete-heading",
          type: "deleteBlock",
          blockId: commentedBlockId,
        },
        edit,
      ],
    } as const satisfies FolioDocumentOperationBatch;

    for (const dryRun of [true, false]) {
      const result = reviewer.applyDocumentOperations(
        { ...batch, dryRun },
        { undefinedReferences: "refuse" },
      );
      expect(result.status).toBe(dryRun ? "previewed" : "rejected");
      expect(result.applied).toEqual([]);
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
        { id: "delete-heading", reason: "invalidResult" },
        { id: edit.id, reason: "invalidResult" },
      ]);
      expect(result.skipped.at(0)?.message).toMatch(/Accept All.*comment/u);
      expect(result.receipts).toEqual([]);
      expect(result.undoHandle).toBeNull();
      expect(reviewer.toDocument()).toEqual(original);
      expect(reviewer.getComments()).toEqual(comments);
      expect(reviewer.getChanges()).toEqual([]);
    }
  });

  test(`${session} admits a reversible tracked edit beside a commented block`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
      session,
    });
    const original = reviewer.toDocument().package.document.content;
    const comments = reviewer.getComments();
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [edit],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.status).toBe("committed");
    expect(result.applied.map(({ id }) => id)).toEqual([edit.id]);
    expect(result.skipped).toEqual([]);

    const pending = await reviewer.toBuffer();
    const accepted = await FolioDocxReviewer.fromBuffer(pending, { session });
    expect(accepted.acceptAll().skipped).toEqual([]);
    expect(accepted.getContent().map(({ text }) => text)).toEqual([
      commentedText,
      "Intervening paragraph.",
      replacementText,
    ]);
    expect(accepted.getComments()).toEqual(comments);

    const rejected = await FolioDocxReviewer.fromBuffer(pending, { session });
    expect(rejected.rejectAll().skipped).toEqual([]);
    expect(rejected.toDocument().package.document.content).toEqual(original);
    expect(rejected.getComments()).toEqual(comments);
  });
}
