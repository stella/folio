export { DocxEditor } from "./components/DocxEditor";
export type {
  DocxEditorCollaboration,
  DocxEditorProps,
  DocxEditorRef,
} from "./components/DocxEditor.props";
export type { EditorMode } from "./components/hooks/useEditorMode";
export type { ColorPreset, FolioButtonProps, FolioUIComponents, OutlineItem } from "./ui/folio-ui";
// Consumers rendering folio chrome (e.g. FormattingBar) outside DocxEditor wrap
// it in this provider to inject their own UI components.
export { FolioUIProvider } from "./ui/folio-ui";
export { FormattingBar, type FormattingBarProps } from "./components/FormattingBar";
export { ZoomControl, type ZoomControlProps, type ZoomLevel } from "./components/ui/ZoomControl";
export {
  clampZoom,
  formatZoom,
  parseZoom,
  useWheelZoom,
  ZOOM_PRESETS,
  type UseWheelZoomOptions,
  type UseWheelZoomReturn,
} from "./hooks/useWheelZoom";
export {
  createEmptyDocument,
  type CreateEmptyDocumentOptions,
} from "@stll/folio-core/utils/createDocument";
export { createDocx } from "@stll/folio-core/docx/rezip";
export type { Document } from "@stll/folio-core/types/document";
export type { DocxCompatibility } from "@stll/folio-core/docx/compatibility";
export {
  deriveBlockId,
  getFolioParaIdFromBlockId,
  isFolioBlockId,
  isSequentialFolioBlockId,
  type DeriveBlockIdInput,
  type FolioBlockId,
} from "@stll/folio-core/types/block-id";

// AI suggestion primitives — types, conflict resolution, apply, and
// the prosemirror decoration plugin. The bar/panel UI itself lives in
// apps/web; folio only ships the headless pieces.
export {
  DEFAULT_AI_SUGGESTION_PRESETS,
  type AICitation,
  type AICitationSource,
  type AIChatMode,
  type AISuggestion,
  type AISuggestionApplyMode,
  type AISuggestionPreset,
  type AISuggestionSeverity,
  type AISuggestionStatus,
  type AIBarStatus,
  type AIGenerateInput,
} from "@stll/folio-core/ai-suggestions/types";
export { applySuggestions, type ApplyResult } from "@stll/folio-core/ai-suggestions/apply";
export {
  applyFolioAIEditOperations,
  createFolioAIEditSnapshot,
  diffWordSegments,
  hashFolioAIBlockText,
  normalizeFolioAIBlockText,
  type WordDiffSegment,
  type FolioAIBlock,
  type FolioAIBlockAnchor,
  type FolioAIBlockKind,
  type FolioAIBlockPreviewRun,
  type FolioAIComment,
  type FolioAIEditAppliedOperation,
  type FolioAIEditApplyMode,
  type FolioAIEditApplyResult,
  type FolioAIEditOperation,
  type FolioAIEditReviewMeta,
  type FolioAIEditSeverity,
  type FolioAIEditSkipReason,
  type FolioAIEditSkippedOperation,
  type FolioAIEditSnapshot,
  type FolioAISignatureParty,
} from "@stll/folio-core/ai-edits";
export {
  resolveSuggestionAnchor,
  isSuggestionStale,
  type ResolvedAnchor,
} from "@stll/folio-core/ai-suggestions/conflict";
export {
  buildPositionalText,
  type PositionalText,
} from "@stll/folio-core/ai-suggestions/text-positions";
export {
  setAISuggestionsMeta,
  setFocusedSuggestionMeta,
} from "@stll/folio-core/prosemirror/plugins/aiSuggestionDecorations";
export { scrollFolioPositionIntoView } from "@stll/folio-core/paged-layout/scrollToPmPosition";
export {
  getFolioCaretViewportRect,
  getFolioSelectionViewportRect,
} from "@stll/folio-core/paged-layout/selectionViewportRect";
export {
  createAICitationDecorationsPlugin,
  setAICitationsMeta,
  setActiveCitationMeta,
  type AICitationRange,
} from "@stll/folio-core/prosemirror/plugins/aiCitationDecorations";
export {
  anonymizationDecorationsKey,
  getAnonymizationMatches,
  setAnonymizationTermsMeta,
  type AnonymizationMatch,
  type AnonymizationTerm,
} from "@stll/folio-core/prosemirror/plugins/anonymizationDecorations";
export {
  getTemplateDirectives,
  scanDirectives,
  type DirectiveKind,
  type DirectiveRange,
} from "@stll/folio-core/prosemirror/plugins/templateDirectives";
export {
  setTemplatePreviewValues,
  type TemplatePreviewSpan,
  type TemplatePreviewValue,
  type TemplatePreviewValues,
} from "@stll/folio-core/prosemirror/plugins/templatePreviewValues";
export {
  clearTemplateSlashMenu,
  consumeTemplateSlashQuery,
  getTemplateSlashMenu,
  resetTemplateSlashQuery,
  templateSlashMenuKey,
  type TemplateSlashMenuKeyAction,
  type TemplateSlashMenuState,
} from "@stll/folio-core/prosemirror/plugins/templateSlashMenu";
export {
  acceptAutocompleteSuggestion,
  acceptAutocompleteWord,
  appendAutocompleteToken,
  autocompleteSuggestionKey,
  autocompleteSuggestionPlugin,
  clearAutocompleteSuggestion,
  DEFAULT_AUTOCOMPLETE_DEAD_ZONE_NODES,
  finishAutocompleteSuggestion,
  getAutocompleteSuggestion,
  shouldTriggerAutocomplete,
  startAutocompleteSuggestion,
  type AcceptAutocompleteResult,
  type AutocompleteSuggestionPluginOptions,
  type AutocompleteSuggestionState,
  type AutocompleteSuggestionStatus,
  type AutocompleteTriggerCheck,
  type AutocompleteTriggerOptions,
  type AutocompleteTriggerSkipReason,
} from "@stll/folio-core/prosemirror/plugins/autocompleteSuggestion";
export {
  AutocompleteCaretOverlay,
  type AutocompleteCaretOverlayProps,
  type AutocompleteCaretRect,
} from "./paged-editor/AutocompleteCaretOverlay";

// DOCX-document ↔ Markdown bridge (also available at `@stll/folio/markdown`).
export {
  fromMarkdown,
  toMarkdown,
  toMarkdownResult,
  type ImageMeta,
  type ImageRef,
  type MarkdownOptions,
  type MarkdownResult,
} from "@stll/folio-core/markdown";
