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
  mode: "tracked-changes",
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

/** A tracked edit that deletes a note, drawing or field shows it deleted. */
export const applyTrackedProtectedEdit = (options: ProtectedEditOptions) => {
  const result = edit(options, "tracked-changes");
  assert.deepEqual(result.skipped, [], JSON.stringify(result));
  assert.equal(result.applied.length, 1, JSON.stringify(result));
  assert.ok(options.reviewer.getChanges().length > 0);
};
