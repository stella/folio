/**
 * `@stll/folio-vue`
 *
 * Root entry for the Vue 3 editor. Mirrors the React adapter's public surface
 * (`packages/react/src/index.ts`): every framework-neutral symbol re-exports the
 * SAME `@stll/folio-core/*` binding React ships, and the framework-specific
 * pieces (the editor component, `renderAsync`, i18n runtime) resolve to their Vue
 * equivalents. A parity CI gate diffs this against the React index.
 *
 * Advanced surfaces stay public through explicit subpaths:
 * `./ui`, `./composables`, `./dialogs`, `./plugin-api`, `./styles`, `./messages`.
 *
 * @packageDocumentation
 * @public
 */

// ─── Editor component + imperative API (Vue-specific) ───────────────────────
export { default as DocxEditor } from "./components/DocxEditor.vue";
// Standalone chrome components (mirror React's index): a host can render these
// outside DocxEditor. Both are controlled — they own no editor view.
export { default as FormattingBar } from "./components/FormattingBar.vue";
export type { FormattingBarProps } from "./components/FormattingBar.types";
export {
  FindReplaceDialog,
  FootnotePropertiesDialog,
  HyperlinkDialog,
  ImagePositionDialog,
  ImagePropertiesDialog,
  InsertImageDialog,
  InsertSymbolDialog,
  InsertTableDialog,
  PageSetupDialog,
  PasteSpecialDialog,
  SplitCellDialog,
  TablePropertiesDialog,
  WatermarkDialog,
} from "./components/dialogs";
export { default as ZoomControl } from "./components/ui/ZoomControl.vue";
export type { ZoomControlProps, ZoomLevel } from "./components/ui/ZoomControl.types";
export {
  renderAsync,
  type DocxEditorHandle,
  type EditorHandle,
  type RenderAsyncOptions,
} from "./renderAsync";
export type {
  DocxEditorCollaboration,
  DocxEditorProps,
  DocxEditorRef,
  EditorMode,
} from "./components/DocxEditor/types";
export type { ScrollToParaIdOptions } from "@stll/folio-core/paged-layout/paragraphFlash";
// Custom-font prop types: `fontFamilies` (picker list) + `fonts` (FontFace registration).
export type { FontOption } from "./utils/fontOptions";
export type { FontDefinition } from "./components/DocxEditor/types";
// Consumer UI-injection contract (mirrors React's `FolioUIComponents`). The Vue
// adapter injects overrides through `DocxEditor`'s `components` prop; the chrome
// resolves primitives via the FolioUI provider. `FolioUIProvider` is a Vue
// component (defined in the `.ts` module so Nuxt's `tsc` typecheck resolves the
// named export). See the parity report for the primitives without a Vue default.
export { FolioUIProvider } from "./ui/folio-ui";
export type { ColorPreset, FolioButtonProps, FolioUIComponents, OutlineItem } from "./ui/folio-ui";

// i18n runtime (Vue-specific — React consumes `use-intl` directly). Locale-string
// types live in the shared catalog; import them from `@stll/folio-core` if needed.
export { defaultLocale, i18nPlugin, provideLocale, useTranslation } from "./i18n";

// Wheel/keyboard zoom composable (Vue equivalent of React's `useWheelZoom` hook).
export { useWheelZoom } from "./composables/useWheelZoom";

// View-level insert operations — consumers wire these into the toolbar's
// `onInsert*` handlers (the Insert group renders only when a handler is passed).
export {
  insertImageFromFile,
  insertPageBreakInView,
  insertTableInView,
  insertTableOfContentsInView,
} from "@stll/folio-core/prosemirror";
export {
  getDocumentWatermark,
  setDocumentWatermark,
  type PictureWatermark,
  type TextWatermark,
  type Watermark,
} from "@stll/folio-core/watermark";

// ─── Framework-neutral core re-exports (identical to React's index) ─────────
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

// DOCX-document ↔ Markdown bridge (also available at `@stll/folio-vue/markdown`).
export {
  fromMarkdown,
  toMarkdown,
  toMarkdownResult,
  type ImageMeta,
  type ImageRef,
  type MarkdownOptions,
  type MarkdownResult,
} from "@stll/folio-core/markdown";
