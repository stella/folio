# @stll/folio-core

## 0.3.1

### Patch Changes

- [#84](https://github.com/stella/folio/pull/84) [`135116d`](https://github.com/stella/folio/commit/135116d9f09fb322b9f95754791bc3b54f637c0d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add an Eigenpal core compatibility entrypoint and support the legacy Google Fonts toggle.

## 0.3.0

### Minor Changes

- [#65](https://github.com/stella/folio/pull/65) [`ca9d64b`](https://github.com/stella/folio/commit/ca9d64bd36bbce78b3dad9aab72092cceebc4919) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `@stll/folio-agents`: a framework-neutral LLM tool layer over folio-core's AI-edits engine — provider-neutral function-calling tool definitions (`read_document`, `find_text`, `read_comments`, `read_changes`, `add_comment`, `suggest_changes`, `reply_comment`, `resolve_comment`, plus live-editor capability tools), an `executeFolioToolCall` executor, Anthropic/OpenAI schema mappers, and bridges for both the headless `FolioDocxReviewer` and a live `DocxEditorRef`. In core, `FolioDocxReviewer` gains `resolveComment(commentId, { resolved? })`, which round-trips `w15:done` through `getComments` and `toBuffer`. It also gains a document version-diff engine — `compareDocxVersions` aligns two `.docx` buffers block by block (stable ids, then text LCS, then positional fallback) into added/deleted/modified changes, and `formatVersionDiffForLLM` renders the result as compact, deterministic text for a model prompt.

- [#81](https://github.com/stella/folio/pull/81) [`439a8ea`](https://github.com/stella/folio/commit/439a8ea12227aafc49e99693755fc03577d4b54c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Broaden the AI-edits read surface for live-editor reuse. Core gains pure `getTrackedChangesFromDoc` / `getCommentAnchorsFromDoc` readers and a `buildAnnotatedBlockText` redline renderer over any ProseMirror `doc`, a `getContentAsText({ annotated })` option that inlines tracked changes and comment anchors as `<ins>`/`<del>`/`<comment>` tags, and a read-only `getNotesAsText` surfacing header/footer and footnote/endnote text. The React and Vue `DocxEditorRef` gain a matching read surface — `getTrackedChanges`, `getCommentAnchors`, `getSelectionText`, and `getPageText(page)` — so a live-editor agent tool can read the current document state without a fresh AI-edit snapshot. `@stll/folio-agents` now exports `parseSuggestChangesInput` / `parseAddCommentInput` — the `suggest_changes` / `add_comment` argument-validation rules factored out of `executeFolioToolCall`, for hosts with their own review-queue UX — and its live-editor bridge now reads real tracked changes and comment anchors from that new `DocxEditorRef` surface instead of reporting them as unavailable.

- [#80](https://github.com/stella/folio/pull/80) [`ce39713`](https://github.com/stella/folio/commit/ce397130771ec8f1271b7d0cfde9e6dc0367e857) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add the UI catalog keys the Vue adapter references but the flat `folio` catalog
  was missing, so components render readable English labels instead of raw
  key-paths (and stop logging `IntlError: MISSING_MESSAGE`).

  Adds 320 keys across the `alignment`, `colorPicker`, `common`, `contextMenu`,
  `dialogs` (hyperlink / insertTable / footnoteProperties / tableProperties /
  pageSetup / findReplace / insertSymbol / imageProperties / imagePosition),
  `editor`, `font`, `formattingBar`, `imageOverlay`, `imageTransform`, `imageWrap`,
  `lineSpacing`, `revisions`, `styles`, `table` (+ `table.styles`), `tableAdvanced`,
  `toolbar`, `trackedChanges`, `viewer`, and `zoom` families, plus flat
  `decreaseFontSize` / `increaseFontSize`. English source values are placeholders in
  the 16 non-English locales (grandfathered in the i18n-check baseline); real
  translations are a separate concern.

  The existing flat `fontSize` label key blocks a `fontSize.*` namespace, so the
  Vue toolbar's font-size step buttons resolve `decreaseFontSize` /
  `increaseFontSize` instead.

## 0.2.0

### Minor Changes

- [#54](https://github.com/stella/folio/pull/54) [`444c11b`](https://github.com/stella/folio/commit/444c11b6c165864bbea59e4ed54c498d1b6fa02e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Readable template directive rails: gutter rails indent by nesting depth
  within a margin-confined budget, colour by directive kind (loop violet,
  condition teal, theme-aware), render quiet by default and emphasize the
  block containing the caret or under the pointer, split into per-page
  segments so no rail crosses a page header, footer, or gap, and pair
  opener/closer chips on hover with a hint of what a bare closer closes.
  Adds pure helpers `computeBlockDepths` (core) and gutter geometry
  measurement to the paged-layout range projection.

- [#57](https://github.com/stella/folio/pull/57) [`c40cb00`](https://github.com/stella/folio/commit/c40cb00390db3abf109c3764645a87451fb6a249) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `@stll/folio-vue`, a Vue 3 editor adapter over `@stll/folio-core` that tracks
  the `@stll/folio-react` editor contract (`DocxEditor`, `DocxEditorProps`,
  `DocxEditorRef`, `renderAsync`).

  To share one framework-neutral base across adapters, the folio UI translation
  catalog moves into `@stll/folio-core/i18n/messages` (the React `messages` subpath
  re-exports it, so `@stll/folio-react/messages` is unchanged). Core also gains the
  helpers the adapters build on: `ClipboardManager`, `AutoSaveManager`,
  `resolveColorToHex`, the `WrapType` union, `docx` + `prosemirror/extensions`
  barrels, and a set of ported editor-engine helpers (comment/section-break/table/
  image commands, tracked-change extraction, visual-line navigation, image layout).

### Patch Changes

- [#56](https://github.com/stella/folio/pull/56) [`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - fix(core): use a CJK font's line height for CJK text in non-CJK-font runs

- [#56](https://github.com/stella/folio/pull/56) [`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - fix(core): derive single-line height from real font hhea metrics

  Single-line height per font is now derived from the font's real `hhea`
  metrics `(ascent + |descent| + lineGap) / unitsPerEm` — the value Word uses —
  instead of hand-transcribed constants, several of which dropped the line gap
  or were otherwise wrong. Corrects 9 fonts (measured against Word): Palatino
  Linotype (was 31% short), Book Antiqua (17%), Cambria (8%), Century Gothic
  (6%), Times New Roman (4%), Arial (3%), Trebuchet MS (2%), Consolas (1%), and
  Lucida Console (14% tall). A shared derivation and a discriminated `hhea` /
  `legacy` representation make it structurally impossible to silently drop a
  term again. CJK fonts are unchanged (their line height is an East-Asian
  layout concern, not a run-font hhea ratio).

- [#55](https://github.com/stella/folio/pull/55) [`e44ba30`](https://github.com/stella/folio/commit/e44ba300c0978b9e2836bafcfe84cd06491e87bf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - fix(core): apply table style paragraph spacing to cell paragraphs

  Table cells now inherit paragraph spacing (space-after, line spacing,
  contextual spacing) from the enclosing table style's `w:pPr` — and from the
  applicable `w:tblStylePr` conditional region (first row, banding, etc.) —
  instead of falling through to `docDefaults`. Per ECMA-376 §17.7.2 this table
  style layer sits between docDefaults and the cell paragraph's own style
  chain/direct formatting, so an explicit paragraph style or direct spacing on
  the paragraph still wins. Previously, cell paragraphs with neither a
  `w:pStyle` nor a direct `w:pPr` picked up the document's default paragraph
  spacing (e.g. Word's default ~10pt space-after and 1.15x line spacing)
  instead of the table style's typically compact spacing (e.g. `TableGrid`'s
  zero space-after / single line), inflating table row heights.

## 0.1.3

### Patch Changes

- Fix DOCX parsing and layout edge cases: smartTag-wrapped runs, EMF header
  previews, percent-suffixed table widths, CJK line breaks, off-page
  header/footer floats, and right-tab trailing width reservation.
