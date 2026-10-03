export { noteUsesCustomMark } from "./noteMarks";
export { sectionPropertiesAt } from "./storyLifecycle";
export { storyBody, findStoryBody, documentStories, sameStory } from "./stories";
/**
 * `@stll/docx-core/ops`: document operations over the typed model.
 *
 * The module depends on the model alone, so an editor, a server and a
 * sequencer apply operations with one implementation.
 */

export { captureDocumentOp } from "./wire";
export { combineEdits } from "./edits";
export {
  applyDocumentOpEnvelope,
  applyDocumentOp,
  applyDocumentOps,
  type AppliedDocumentOp,
} from "./apply";
export { formattingEquals, isTextOnlyRun, runsMergeable, mergeRunContent } from "./runMerge";
export {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  paragraphVisibleText,
  physicalOffsetAtVisibleOffset,
  editorParagraphGroups,
  type EditorIntent,
  type EditorIntentMode,
} from "./editorIntent";
export { DocumentOpsContractError, normalizeForOps, validateOpsDocument } from "./contract";
export { OBJECT_REPLACEMENT_CHARACTER, paragraphLength, paragraphLogicalText } from "./offsets";
export { planTrackedReplace, type PlanTrackedReplaceOptions } from "./rangeReplacement";
export {
  BATCH_REJECTION_REASONS,
  BATCH_WIRE_OP_TYPES,
  MAX_BATCH_WIRE_BYTES,
  BatchRejection,
  parseDocumentBatch,
  validateDocumentBatch,
  validateSequencedBatch,
  type BatchRejectionReason,
  type DocumentBatch,
  type SequencedBatch,
} from "./sequencing/envelope";
export { transformBatch } from "./sequencing/transform";
export { createSequencer } from "./sequencing/sequencer";
export { createClient } from "./sequencing/client";
export { planTrackedDeletion, type PlanTrackedDeletionOptions, revisionIdDemand } from "./plan";
export {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
export {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  PROPERTY_REVIEW_POLICIES,
  EMPTY_PROPERTY_SETS,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  PARAGRAPH_MARK_FORMATTING_KEYS,
  REVISION_DECISIONS,
  SPLIT_HALVES,
  toOpEnvelope,
  type HeaderFooterStory,
  type NoteStory,
  type CreateHeaderFooterOp,
  type RemoveHeaderFooterOp,
  type AddNoteOp,
  type RemoveNoteOp,
  type SetSectionPropsOp,
  type RestoreStoryPartsOp,
  type BlockInsertionPoint,
  type CreateNumberingInstanceOp,
  type InsertBlocksOp,
  type DeleteBlocksOp,
  type DeleteNumberingInstanceOp,
  type InsertTableOp,
  type DeleteTableOp,
  type SetContainerBlocksOp,
  type DeleteRangeOp,
  type DeleteRowOp,
  type DocumentOp,
  type DocumentOpEnvelope,
  type DocumentOpType,
  type EmptyPropertySet,
  type FormattingPatch,
  type InlineSlice,
  type InsertContentOp,
  type InsertedRunProps,
  type InsertTextOp,
  type InsertRowOp,
  type JoinBlocksOp,
  type JoinInlineOp,
  type NewIds,
  type NumberingPartState,
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
  type SetTableRowsOp,
  type SectionPropertiesState,
  type SectionMapState,
  type SectionViewEntry,
  type SectionViewState,
  type SectionEndpoint,
  type SetSectionEndpointOp,
  type SplitBlockOp,
  type SplitHalf,
  type SplitInlineOp,
  type SplitParagraphFields,
  type TextPosition,
  type TouchedBlocks,
} from "./types";

export { packageParagraphIds } from "./ids";
