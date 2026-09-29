/**
 * `@stll/docx-core/ops`: document operations over the typed model.
 *
 * The module depends on the model alone, so an editor, a server and a
 * sequencer apply operations with one implementation.
 */

export { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "./apply";
export { DocumentOpsContractError, normalizeForOps, validateOpsDocument } from "./contract";
export { OBJECT_REPLACEMENT_CHARACTER, paragraphLength, paragraphLogicalText } from "./offsets";
export { planTrackedDeletion, type PlanTrackedDeletionOptions, revisionIdDemand } from "./plan";
export {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
export {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  EMPTY_PROPERTY_SETS,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  PARAGRAPH_MARK_FORMATTING_KEYS,
  REVISION_DECISIONS,
  SPLIT_HALVES,
  toOpEnvelope,
  type DeleteRangeOp,
  type DocumentOp,
  type DocumentOpEnvelope,
  type DocumentOpType,
  type EmptyPropertySet,
  type FormattingPatch,
  type InlineSlice,
  type InsertContentOp,
  type InsertedRunProps,
  type InsertTextOp,
  type JoinBlocksOp,
  type JoinInlineOp,
  type NewIds,
  type OpStory,
  type ParagraphPropsPatch,
  type ParagraphReviewFields,
  type ReplaceBlocksOp,
  type ReplaceInlineOp,
  type ResolveRevisionOp,
  type RevisionDecision,
  type RevisionStamp,
  type RunPropsPatch,
  type SetParagraphPropsOp,
  type SetParagraphReviewOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  type SplitHalf,
  type SplitInlineOp,
  type SplitParagraphFields,
  type TextPosition,
  type TouchedBlocks,
} from "./types";
