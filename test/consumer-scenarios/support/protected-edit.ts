import assert from "node:assert/strict";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";

import { openReviewer, packDocument } from "./documents.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Story = Parameters<Reviewer["applyDocumentOperationsToStory"]>[0]["story"];

type ProtectedEditOptions = {
  reviewer: Reviewer;
  operation: FolioDocumentOperation;
  kind: "comment" | "note";
  story?: Story;
};

const edit = (
  { reviewer, operation, story }: ProtectedEditOptions,
  mode: "tracked-changes" | "suggested",
) => {
  const batch = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [operation],
  } as const;
  return story === undefined
    ? reviewer.applyDocumentOperations(batch, { undefinedReferences: "refuse" })
    : reviewer.applyDocumentOperationsToStory({
        story,
        batch,
        undefinedReferences: "refuse",
      });
};

export const assertProtectedEditRefused = (options: ProtectedEditOptions) => {
  const { reviewer, operation, kind } = options;
  const before = reviewer.toDocument();
  const result = edit(options, "tracked-changes");
  assert.equal(result.status, "rejected", JSON.stringify(result));
  assert.deepEqual(result.applied, []);
  assert.deepEqual(result.skipped, [
    {
      id: operation.id,
      reason: "invalidResult",
      message: `Accept All would lose a ${kind}.`,
    },
  ]);
  assert.deepEqual(result.receipts, []);
  assert.equal(result.undoHandle, null);
  assert.deepEqual(reviewer.toDocument(), before);
};

export const stageReviewedProtectedEdit = (options: ProtectedEditOptions) => {
  assertProtectedEditRefused(options);
  const result = edit(options, "suggested");
  assert.deepEqual(result.skipped, []);
  const suggestionId = result.applied.at(0)?.suggestionId;
  assert.ok(suggestionId, JSON.stringify(result));
  assert.equal(options.reviewer.acceptSuggestion(suggestionId), true);
};

/** Import pending row deletions for review tests that need an authored table removal. */
export const importTableDeletion = async (
  reviewer: Reviewer,
  blockId: string,
) => {
  const document = structuredClone(reviewer.toDocument());
  const table = document.package.document.content.find(
    (block) =>
      block.type === "table" &&
      block.rows.some(({ cells }) =>
        cells.some(({ content }) =>
          content.some(
            (paragraph) =>
              paragraph.type === "paragraph" && paragraph.paraId === blockId,
          ),
        ),
      ),
  );
  assert.ok(table?.type === "table", "Expected the fixture's top-level table");
  const info = { id: 100_000, author: "Reviewer" };
  assert.ok(reviewer.getChanges().every(({ id }) => id !== info.id));
  for (const row of table.rows) {
    if (row.structuralChange?.type === "tableRowDeletion") continue;
    assert.equal(row.structuralChange, undefined);
    row.structuralChange = { type: "tableRowDeletion", info };
  }
  const imported = await openReviewer(await packDocument(document));
  assert.ok(
    imported
      .getChanges()
      .some(({ id, type }) => id === info.id && type === "rowDeleted"),
    "Expected pending row deletion after importing the fixture",
  );
  return imported;
};
