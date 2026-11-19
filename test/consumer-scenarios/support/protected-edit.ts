import assert from "node:assert/strict";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";

import type { openReviewer } from "./documents.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Story = Parameters<Reviewer["applyDocumentOperationsToStory"]>[0]["story"];

type ProtectedEditOptions = {
  reviewer: Reviewer;
  operation: FolioDocumentOperation;
  kind: "note";
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

const assertProtectedEditRefused = (options: ProtectedEditOptions) => {
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
