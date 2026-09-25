/**
 * `@stll/docx-core/ops`: document operations over the typed model.
 *
 * The module depends on the model alone, so an editor, a server and a
 * sequencer apply operations with one implementation.
 */

export { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "./apply";
export { DocumentOpsContractError, normalizeForOps, validateOpsDocument } from "./contract";
export { OBJECT_REPLACEMENT_CHARACTER, paragraphLength, paragraphLogicalText } from "./offsets";
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
  type ReplaceBlocksOp,
  type RunPropsPatch,
  type SetParagraphPropsOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  type SplitInlineOp,
  type SplitParagraphFields,
  type TextPosition,
  type TouchedBlocks,
} from "./types";
