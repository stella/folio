/**
 * Server-only entry for `@stll/folio`.
 *
 * Type re-exports (folio block data shapes) plus the small set of
 * DOM-free runtime helpers servers need to PRODUCE ids that round-
 * trip through the editor — sharing `deriveBlockId` here is what
 * keeps the server DOCX parser and the in-browser snapshot from
 * minting incompatible block ids.
 *
 * Anything DOM-dependent stays on the main `@stll/folio` entry.
 */
export type {
  FolioAIBlock,
  FolioAIBlockAnchor,
  FolioAIBlockKind,
  FolioAIBlockPreviewRun,
  FolioAIEditSnapshot,
  FolioAIInlineFormatting,
  FolioAITextRangeHandle,
} from "./ai-edits/types";
export { createFolioAITextRangeHandle } from "./ai-edits/snapshot";
export {
  deriveBlockId,
  getFolioParaIdFromBlockId,
  isFolioBlockId,
  isSequentialFolioBlockId,
  type DeriveBlockIdInput,
  type FolioBlockId,
} from "./types/block-id";
export { createDocx } from "./docx/rezip";
export { replyToComment, type CreateCommentReplyInput } from "./docx/replyToComment";
export { createEmptyDocument, type CreateEmptyDocumentOptions } from "./utils/createDocument";
export {
  FolioDocxReviewer,
  applyFolioAIEditsToBuffer,
  type ApplyFolioAIEditsToBufferOptions,
  type ApplyFolioAIEditsToBufferResult,
  type FolioApplyOperationsOptions,
  type FolioApplyDocumentOperationsOptions,
  type FolioDocxReviewerOptions,
  type FolioReviewChange,
  type FolioReviewChangeFilter,
  type FolioReviewChangeKind,
  type FolioReviewComment,
  type FolioReviewCommentFilter,
  type FolioReviewCommentReply,
  type FolioReviewReplyInput,
} from "./ai-edits/headless";
export type {
  FolioAIComment,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditPrecondition,
} from "./ai-edits/types";
export {
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_BATCH_MODES,
  FOLIO_DOCUMENT_OPERATION_MODES,
  FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE,
  FOLIO_DOCUMENT_OPERATION_PRECONDITIONS,
  FOLIO_DOCUMENT_OPERATION_STORIES,
  FOLIO_DOCUMENT_OPERATION_TYPES,
  getFolioDocumentOperationCapabilities,
  InvalidFolioDocumentOperationBatchError,
  isSupportedFolioDocumentOperationVersion,
  isFolioDocumentOperationModeSupported,
  parseFolioDocumentOperationBatch,
  UnsupportedFolioDocumentOperationVersionError,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationCapabilities,
  type FolioDocumentOperationMode,
  type FolioDocumentOperationPrecondition,
  type FolioDocumentOperationType,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationStatus,
} from "./document-operations";
