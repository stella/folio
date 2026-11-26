/** Compatibility exports for the deep comments command path. */
export {
  addCommentMark,
  findChangeAtPosition,
  findNextChange,
  findParagraphBoundaryChangeAtPosition,
  findPreviousChange,
  type ParagraphBoundaryChange,
  removeCommentMark,
  trackedChangeEntryRevisionIds,
} from "../projectionReview";
export { isPPrMarkAttr, paragraphMarkWasAdded } from "../revisionCarriers";
export {
  findAIEditRevisionRange,
  isTableCellRevisionAttr,
  isTableRowRevisionAttr,
  type TableRowRevisionAttr,
} from "../revisionRangeQueries";
export {
  collectSuggestions,
  findSuggestionRange,
  getSuggestions,
  isSuggestionMark,
  mergeSegments,
  pendingSuggestionRevisionIds,
  readStructuralSuggestion,
  readSuggestedMarker,
  readSuggestedParagraphMark,
  readSuggestedParagraphPropertyChange,
  suggestedCommentOf,
  suggestionCommentIds,
  suggestionIdOfMarkRecord,
  suggestionIdOfRevision,
  suggestionOwnersByRevision,
  type FolioSuggestion,
  type SuggestionAppliedAs,
  type SuggestionKind,
} from "../suggestionQueries";
