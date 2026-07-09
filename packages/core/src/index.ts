// Headless public API for `@stll/folio/core`.
//
// Everything reachable from this entry is framework-neutral (the
// `no-react-in-core` lint rule enforces it), so non-React adapters (a Vue
// adapter, a Tauri shell, a Rust core) can build on the editor core without
// pulling React or @stll/ui into their import graph. The root `@stll/folio`
// entry re-exports all of this and adds the React components on top.

export { createEmptyDocument, type CreateEmptyDocumentOptions } from "./utils/createDocument";
export { createDocx } from "./docx/rezip";
export type { Document } from "./types/document";
export type { DocxCompatibility } from "./docx/compatibility";
export {
  deriveBlockId,
  getFolioParaIdFromBlockId,
  isFolioBlockId,
  isSequentialFolioBlockId,
  type DeriveBlockIdInput,
  type FolioBlockId,
} from "./types/block-id";

// AI suggestion primitives — types, conflict resolution, apply, and the
// prosemirror decoration plugin. The bar/panel UI itself lives in apps/web;
// folio only ships the headless pieces.
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
} from "./ai-suggestions/types";
export { applySuggestions, type ApplyResult } from "./ai-suggestions/apply";
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
  applyFolioDocumentOperations,
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_MODES,
  FOLIO_DOCUMENT_OPERATION_STORIES,
  FOLIO_DOCUMENT_OPERATION_TYPES,
  getFolioDocumentOperationCapabilities,
  isSupportedFolioDocumentOperationVersion,
  UnsupportedFolioDocumentOperationVersionError,
  type ApplyFolioDocumentOperationsOptions,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationCapabilities,
  type FolioDocumentOperationMode,
  type FolioDocumentOperationResult,
} from "./ai-edits";
export {
  resolveSuggestionAnchor,
  isSuggestionStale,
  type ResolvedAnchor,
} from "./ai-suggestions/conflict";
export { buildPositionalText, type PositionalText } from "./ai-suggestions/text-positions";
export {
  setAISuggestionsMeta,
  setFocusedSuggestionMeta,
} from "./prosemirror/plugins/aiSuggestionDecorations";
export { scrollFolioPositionIntoView } from "./paged-layout/scrollToPmPosition";
export {
  getFolioCaretViewportRect,
  getFolioSelectionViewportRect,
} from "./paged-layout/selectionViewportRect";
export {
  createAICitationDecorationsPlugin,
  setAICitationsMeta,
  setActiveCitationMeta,
  type AICitationRange,
} from "./prosemirror/plugins/aiCitationDecorations";
export {
  anonymizationDecorationsKey,
  getAnonymizationMatches,
  setAnonymizationTermsMeta,
  type AnonymizationMatch,
  type AnonymizationTerm,
} from "./prosemirror/plugins/anonymizationDecorations";
export {
  getTemplateDirectives,
  scanDirectives,
  type DirectiveKind,
  type DirectiveRange,
} from "./prosemirror/plugins/templateDirectives";
export {
  setTemplatePreviewValues,
  type TemplatePreviewSpan,
  type TemplatePreviewValue,
  type TemplatePreviewValues,
} from "./prosemirror/plugins/templatePreviewValues";
export {
  clearTemplateSlashMenu,
  consumeTemplateSlashQuery,
  getTemplateSlashMenu,
  resetTemplateSlashQuery,
  templateSlashMenuKey,
  type TemplateSlashMenuKeyAction,
  type TemplateSlashMenuState,
} from "./prosemirror/plugins/templateSlashMenu";
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
} from "./prosemirror/plugins/autocompleteSuggestion";

// DOCX-document ↔ Markdown bridge (also available at `@stll/folio/markdown`).
export {
  fromMarkdown,
  toMarkdown,
  toMarkdownResult,
  type ImageMeta,
  type ImageRef,
  type MarkdownOptions,
  type MarkdownResult,
} from "./markdown";

// Embedded fonts: de-obfuscate the `word/fonts/*.odttf` binaries a document
// carries so it can render in its authored (embedded-only) fonts. Pure and
// DOM-free; the React editor registers the result as `@font-face`s.
export {
  extractEmbeddedFonts,
  getEmbeddedFontFaces,
  type EmbeddedFont,
  type EmbeddedFontParts,
} from "./fonts/embeddedFonts";
export { getGoogleFontsEnabled, setGoogleFontsEnabled } from "./utils/fontResolver";
