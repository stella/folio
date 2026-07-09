import { TaggedError } from "better-result";

import { applyFolioAIEditOperations, type FolioAIEditView } from "./ai-edits/apply";
import type {
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditOperation,
  FolioAIEditSkippedOperation,
  FolioAIEditSnapshot,
} from "./ai-edits/types";

export const FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION = 1 as const;

export const FOLIO_DOCUMENT_OPERATION_TYPES = Object.freeze([
  "replaceInBlock",
  "insertAfterBlock",
  "insertBeforeBlock",
  "replaceBlock",
  "deleteBlock",
  "commentOnBlock",
  "insertSignatureTable",
] as const satisfies readonly FolioAIEditOperation["type"][]);

export const FOLIO_DOCUMENT_OPERATION_MODES = Object.freeze([
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioAIEditApplyMode[]);

export const FOLIO_DOCUMENT_OPERATION_STORIES = Object.freeze(["main"] as const);

export type FolioDocumentOperation = FolioAIEditOperation;
export type FolioDocumentOperationMode = FolioAIEditApplyMode;

export type FolioDocumentOperationCapabilities = {
  readonly version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  readonly operationTypes: typeof FOLIO_DOCUMENT_OPERATION_TYPES;
  readonly modes: typeof FOLIO_DOCUMENT_OPERATION_MODES;
  readonly stories: typeof FOLIO_DOCUMENT_OPERATION_STORIES;
};

const DOCUMENT_OPERATION_CAPABILITIES: FolioDocumentOperationCapabilities = Object.freeze({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  operationTypes: FOLIO_DOCUMENT_OPERATION_TYPES,
  modes: FOLIO_DOCUMENT_OPERATION_MODES,
  stories: FOLIO_DOCUMENT_OPERATION_STORIES,
});

export const getFolioDocumentOperationCapabilities = (): FolioDocumentOperationCapabilities =>
  DOCUMENT_OPERATION_CAPABILITIES;

export const isSupportedFolioDocumentOperationVersion = (
  value: unknown,
): value is typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION =>
  value === FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;

export class UnsupportedFolioDocumentOperationVersionError extends TaggedError(
  "UnsupportedFolioDocumentOperationVersionError",
)<{
  message: string;
  receivedVersion: unknown;
}>() {}

export const assertSupportedFolioDocumentOperationVersion = (
  value: unknown,
): typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION => {
  if (isSupportedFolioDocumentOperationVersion(value)) {
    return value;
  }
  throw new UnsupportedFolioDocumentOperationVersionError({
    message: "Unsupported document operation contract version.",
    receivedVersion: value,
  });
};

export type FolioDocumentOperationBatch = {
  readonly version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  operations: FolioDocumentOperation[];
  mode?: FolioDocumentOperationMode;
};

export type FolioDocumentOperationResult = {
  version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  applied: FolioAIEditAppliedOperation[];
  skipped: FolioAIEditSkippedOperation[];
};

export type ApplyFolioDocumentOperationsOptions = {
  view: FolioAIEditView;
  snapshot: FolioAIEditSnapshot;
  batch: FolioDocumentOperationBatch;
  author?: string;
  createCommentId?: (text: string) => number;
};

export const applyFolioDocumentOperations = ({
  view,
  snapshot,
  batch,
  author,
  createCommentId,
}: ApplyFolioDocumentOperationsOptions): FolioDocumentOperationResult => {
  assertSupportedFolioDocumentOperationVersion(batch.version);
  return {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    ...applyFolioAIEditOperations({
      view,
      snapshot,
      operations: batch.operations,
      mode: batch.mode ?? "tracked-changes",
      ...(author !== undefined && { author }),
      ...(createCommentId !== undefined && { createCommentId }),
    }),
  };
};
