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
  FolioDocumentNavigationTarget,
  FolioDocumentOutline,
  FolioDocumentOutlineEntry,
  FolioDocumentSection,
  FolioDocumentSectionHandle,
  FolioDocumentSectionReadResult,
} from "./ai-edits/types";
export { createFolioAITextRangeHandle } from "./ai-edits/snapshot";
export { getFolioDocumentOutline, readFolioDocumentSection } from "./ai-edits/scoped-reading";
export {
  deriveBlockId,
  getFolioParaIdFromBlockId,
  isFolioBlockId,
  isSequentialFolioBlockId,
  type DeriveBlockIdInput,
  type FolioBlockId,
} from "./types/block-id";
export { createDocx } from "./docx/rezip";
export {
  ensureParaIds,
  EnsureParaIdsError,
  type EnsureParaIdsOptions,
  type EnsureParaIdsResult,
} from "./docx/ensureParaIds";
export { replyToComment, type CreateCommentReplyInput } from "./docx/replyToComment";
export { createEmptyDocument, type CreateEmptyDocumentOptions } from "./utils/createDocument";
export {
  extractDocumentStyleSet,
  extractDocumentStyleSetFromDocx,
  inspectDocumentStyles,
  inspectDocumentStylesFromDocx,
  type DocumentStyleCatalog,
  type DocumentStyleCatalogEntry,
  type ExtractDocumentStyleSetOptions,
} from "./style-sets/extract";
export {
  createStellaStyleDocumentPreset,
  createStellaStyleSet,
  STELLA_STYLE_SET_NAME,
} from "./style-sets/stellaStyle";
export {
  DOCUMENT_PRESET_VERSION,
  DOCUMENT_STYLE_SET_VERSION,
  type DocumentPreset,
  type DocumentStyleSet,
} from "./style-sets/types";
export {
  FOLIO_RESOLVED_REVIEWED_VIEWS,
  FOLIO_REVIEWED_VIEWS,
  FolioDocxReviewer,
  FolioDocumentStoryNotFoundError,
  UnsupportedFolioReviewedViewError,
  applyFolioAIEditsToBuffer,
  isFolioResolvedReviewedView,
  isFolioReviewedView,
  type ApplyFolioAIEditsToBufferOptions,
  type ApplyFolioAIEditsToBufferResult,
  type FolioApplyOperationsOptions,
  type FolioApplyDocumentOperationsOptions,
  type FolioApplyDocumentOperationsToStoryOptions,
  type FolioDocxReviewerOptions,
  type FolioDocumentStory,
  type FolioDocumentStoryHandle,
  type FolioEditableDocumentStoryHandle,
  type FolioReadReviewedStoryOptions,
  type FolioResolveReviewedStoryOptions,
  type FolioResolvedReviewedView,
  type FolioReviewedStory,
  type FolioReviewedView,
  type FolioReviewChange,
  type FolioReviewChangeFilter,
  type FolioReviewChangeKind,
  type FolioReviewComment,
  type FolioReviewCommentFilter,
  type FolioReviewCommentReply,
  type FolioReviewReplyInput,
} from "./ai-edits/headless";
export {
  applyFolioVersionDiffPrivacy,
  compareDocxVersions,
  FOLIO_DOCUMENT_METADATA_PROPERTIES,
  FOLIO_VERSION_COMPARISON_SCOPES,
  FOLIO_VERSION_COMPARISON_PRIVACY_TRANSFORMS,
  InvalidFolioVersionComparisonOptionsError,
  isFolioVersionComparisonScope,
  isFolioVersionComparisonPrivacyTransform,
  type FolioBlockDiff,
  type FolioCompareDocxVersionsOptions,
  type FolioDocumentMetadataProperty,
  type FolioDocumentMetadataValue,
  type FolioFormatProperty,
  type FolioMetadataDiff,
  type FolioStoryDiff,
  type FolioVersionBlockHandle,
  type FolioVersionComparisonScope,
  type FolioVersionComparisonPrivacyTransform,
  type FolioVersionDiff,
  type FolioVersionDiffPrivacyOptions,
  type FolioVersionDiffPrivacyReport,
  type FolioVersionDiffSummaryCounts,
  type FolioVersionDiffSegment,
} from "./version-comparison";
export {
  generateRedlineDocx,
  InvalidGenerateRedlineDocxOptionsError,
  type GenerateRedlineDocxOptions,
  type GenerateRedlineDocxResult,
  type GenerateRedlineUnprocessedStory,
} from "./redline";
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
  getFolioDocumentOperationIssues,
  getFolioDocumentOperationReceipts,
  InvalidFolioDocumentOperationBatchError,
  isSupportedFolioDocumentOperationVersion,
  isFolioDocumentOperationModeSupported,
  parseFolioDocumentOperationBatch,
  UnsupportedFolioDocumentOperationVersionError,
  type FolioDocumentOperation,
  type FolioDocumentOperationAffectedTarget,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationCapabilities,
  type FolioDocumentOperationMode,
  type FolioDocumentOperationIssue,
  type FolioDocumentOperationPrecondition,
  type FolioDocumentOperationRecovery,
  type FolioDocumentOperationReceipt,
  type FolioDocumentOperationType,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationStatus,
  type FolioDocumentOperationStory,
  type FolioDocumentOperationUndoFailureReason,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
} from "./document-operations";
export {
  extractDocxText,
  type DocxParagraphSource,
  type ExtractedDocxParagraph,
  type ExtractedDocxText,
} from "./docx/server/extractDocxText";
export { DocxArchiveError } from "./docx/server/boundedArchive";
export {
  FOLIO_DOCUMENT_PRIVACY_TRANSFORMS,
  FolioDocumentPrivacyArchiveError,
  InvalidFolioDocumentPrivacyOptionsError,
  isFolioDocumentPrivacyTransform,
  rewriteDocxMetadataPrivacy,
  type FolioDocumentPrivacyOptions,
  type FolioDocumentPrivacyReport,
  type FolioDocumentPrivacyTransform,
  type RewriteDocxMetadataPrivacyResult,
} from "./docx/metadataPrivacy";
