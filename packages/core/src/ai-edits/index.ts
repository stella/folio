export { applyFolioAIEditOperations, type FolioAIEditView } from "./apply";
export { buildAnnotatedBlockText } from "./clean-text";
export {
  getCommentAnchorsFromDoc,
  getTrackedChangesFromDoc,
  type FolioCommentAnchor,
  type FolioReviewChange,
  type FolioReviewChangeKind,
} from "./read";
export {
  createFolioAIEditSnapshot,
  hashFolioAIBlockText,
  normalizeFolioAIBlockText,
} from "./snapshot";
export { getFolioParaIdFromBlockId } from "../types/block-id";
export { diffWordSegments } from "./word-diff";
export type { WordDiffSegment } from "./word-diff";
export type {
  FolioAIBlock,
  FolioAIBlockAnchor,
  FolioAIBlockKind,
  FolioAIBlockPreviewRun,
  FolioAIComment,
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditReviewMeta,
  FolioAIEditSeverity,
  FolioAIEditSkipReason,
  FolioAIEditSkippedOperation,
  FolioAIEditSnapshot,
  FolioAISignatureParty,
} from "./types";
export {
  applyFolioDocumentOperations,
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_MODES,
  FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE,
  FOLIO_DOCUMENT_OPERATION_STORIES,
  FOLIO_DOCUMENT_OPERATION_TYPES,
  getFolioDocumentOperationCapabilities,
  InvalidFolioDocumentOperationBatchError,
  isSupportedFolioDocumentOperationVersion,
  isFolioDocumentOperationModeSupported,
  parseFolioDocumentOperationBatch,
  UnsupportedFolioDocumentOperationVersionError,
  type ApplyFolioDocumentOperationsOptions,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationCapabilities,
  type FolioDocumentOperationMode,
  type FolioDocumentOperationType,
  type FolioDocumentOperationResult,
} from "../document-operations";
