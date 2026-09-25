/**
 * `@stll/docx-core/ops`: document operations over the typed model.
 *
 * The module depends on the model alone, so an editor, a server and a
 * sequencer apply operations with one implementation.
 */

export { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "./apply";
export { OBJECT_REPLACEMENT_CHARACTER, paragraphLength, paragraphLogicalText } from "./offsets";
export {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
export {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  type DeleteRangeOp,
  type DocumentOp,
  type DocumentOpType,
  type FormattingPatch,
  type InsertedRunProps,
  type InsertTextOp,
  type JoinBlocksOp,
  type OpStory,
  type ParagraphPropsPatch,
  type ReplaceBlocksOp,
  type RunPropsPatch,
  type SetParagraphPropsOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  type TextPosition,
  type TouchedBlocks,
} from "./types";
