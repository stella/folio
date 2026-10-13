import type { FolioApplyOperationsOptions } from "../../src/ai-edits/headless";
import type { ApplyFolioDocumentOperationsOptions } from "../../src/document-operations";
import type { applyFolioAIEditOperations } from "../../src/ai-edits/apply";

const editing = { undefinedReferences: "refuse" } as const satisfies FolioApplyOperationsOptions;
const comparison = { undefinedReferences: "keep" } as const satisfies FolioApplyOperationsOptions;

// @ts-expect-error reviewer callers must state their reference policy
const missingReviewerPolicy: FolioApplyOperationsOptions = {};
// @ts-expect-error direct document callers must state their reference policy
const missingDocumentPolicy: Pick<ApplyFolioDocumentOperationsOptions, "undefinedReferences"> = {};
// @ts-expect-error direct editor callers must state their reference policy
const missingEditorPolicy: Pick<
  Parameters<typeof applyFolioAIEditOperations>[0],
  "undefinedReferences"
> = {};

export type RequiredReferencePolicyProof =
  | typeof editing
  | typeof comparison
  | typeof missingReviewerPolicy
  | typeof missingDocumentPolicy
  | typeof missingEditorPolicy;
