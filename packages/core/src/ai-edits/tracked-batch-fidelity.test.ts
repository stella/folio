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
const commentedText = "Document heading.";
const commentQuote = "heading";
const originalText = "Original wording";
const replacementText = "Revised wording";

const source = () =>
  paragraphsDocx(
    [
      [
        textRun("Document "),
        { type: "commentRangeStart", id: 1 },
        textRun(commentQuote),
        { type: "commentRangeEnd", id: 1 },
        { type: "commentReference", id: 1 },
        textRun("."),
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

test.each(["legacy", "canonical"] as const)(
  "%s admits a tracked deletion that identifies its retired comment",
  async (session) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
      session,
    });
    const batch = {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "delete-heading",
          type: "replaceInBlock",
          blockId: commentedBlockId,
          find: commentQuote,
          replace: "",
        },
      ],
    } as const satisfies FolioDocumentOperationBatch;
    const before = reviewer.toDocument();
    for (const dryRun of [true, false]) {
      const result = reviewer.applyDocumentOperations(
        { ...batch, dryRun },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied).toHaveLength(1);
      expect(result.receipts.at(0)?.retiredCommentIds).toEqual([1]);
      if (dryRun) expect(reviewer.toDocument()).toEqual(before);
    }
    const pending = await reviewer.toBuffer();
    const accepted = await FolioDocxReviewer.fromBuffer(pending, { session });
    expect(accepted.acceptAll().skipped).toEqual([]);
    expect(accepted.getComments()).toEqual([]);
    const rejected = await FolioDocxReviewer.fromBuffer(pending, { session });
    expect(rejected.rejectAll().skipped).toEqual([]);
    expect(
      rejected.getComments().map(({ anchoredText }) => anchoredText),
    ).toEqual([commentQuote]);
  },
);

test("refuses a tracked batch whose combined deletions remove an unreported comment", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
    session: "legacy",
  });
  const original = reviewer.toDocument();
  const comments = reviewer.getComments();
  expect(comments.map(({ anchoredText }) => anchoredText)).toEqual([
    commentQuote,
  ]);
  const deletions = ["head", "ing"].map((find) => ({
    id: `delete-${find}`,
    type: "replaceInBlock" as const,
    blockId: commentedBlockId,
    find,
    replace: "",
  }));
  const batch = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    atomic: false,
    operations: [...deletions, edit],
  } as const satisfies FolioDocumentOperationBatch;

  for (const dryRun of [true, false]) {
    const result = reviewer.applyDocumentOperations(
      { ...batch, dryRun },
      { undefinedReferences: "refuse" },
    );
    expect(result.status).toBe(dryRun ? "previewed" : "rejected");
    expect(result.applied).toEqual([]);
    expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual(
      batch.operations.map(({ id }) => ({ id, reason: "invalidResult" })),
    );
    expect(result.skipped.at(0)?.message).toMatch(/Accept All.*comment/u);
    expect(result.receipts).toEqual([]);
    expect(result.undoHandle).toBeNull();
    expect(reviewer.toDocument()).toEqual(original);
    expect(reviewer.getComments()).toEqual(comments);
    expect(reviewer.getChanges()).toEqual([]);
  }
});

for (const session of ["legacy", "canonical"] as const) {
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
