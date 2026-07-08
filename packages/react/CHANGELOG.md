# @stll/folio-react

## 0.6.0

### Minor Changes

- [#91](https://github.com/stella/folio/pull/91) [`521879e`](https://github.com/stella/folio/commit/521879e5b3db5c91e8a68dec0de14e31cc964557) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Export React dialogs for watermarks, hyperlinks, insert table/image, paste special, and split cell workflows.

### Patch Changes

- Updated dependencies [[`989999c`](https://github.com/stella/folio/commit/989999cc3629c434a36b16bffee8cba8eb2171b4), [`63731b6`](https://github.com/stella/folio/commit/63731b61e9be9601adf52a6faceaa7cd1ee9fbc4), [`521879e`](https://github.com/stella/folio/commit/521879e5b3db5c91e8a68dec0de14e31cc964557), [`b637aa4`](https://github.com/stella/folio/commit/b637aa44b5f6705affa859c037af96b38df4360d)]:
  - @stll/folio-core@0.4.0

## 0.5.0

### Minor Changes

- [#83](https://github.com/stella/folio/pull/83) [`2965019`](https://github.com/stella/folio/commit/2965019bcb330439baabee5789757827eab757eb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Support React 18 alongside React 19.

- [#84](https://github.com/stella/folio/pull/84) [`135116d`](https://github.com/stella/folio/commit/135116d9f09fb322b9f95754791bc3b54f637c0d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add an Eigenpal compatibility entrypoint for migration codemods.

### Patch Changes

- [#82](https://github.com/stella/folio/pull/82) [`59fac3d`](https://github.com/stella/folio/commit/59fac3d8c57c786db0681e46cb4bb51f4c292992) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fix `FormattingBar` overflow collapse: the secondary control group (font, color, alignment, lists) now collapses into the "More" popover based on live measurement of the toolbar's actual content width (via `ResizeObserver`), instead of a fixed bar-width breakpoint that ignored host `priorityExtra`/`inlineExtra` width and could leave controls scrolled out of view with no visible affordance; the "More" trigger is now rendered outside the scrollable region so it can never scroll away. Also fixes the zoom control and font-size picker truncating their labels, normalizes the alignment/list active-state affordance and icon sizes to match bold/italic/underline.

- Updated dependencies [[`135116d`](https://github.com/stella/folio/commit/135116d9f09fb322b9f95754791bc3b54f637c0d)]:
  - @stll/folio-core@0.3.1

## 0.4.0

### Minor Changes

- [#81](https://github.com/stella/folio/pull/81) [`439a8ea`](https://github.com/stella/folio/commit/439a8ea12227aafc49e99693755fc03577d4b54c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Broaden the AI-edits read surface for live-editor reuse. Core gains pure `getTrackedChangesFromDoc` / `getCommentAnchorsFromDoc` readers and a `buildAnnotatedBlockText` redline renderer over any ProseMirror `doc`, a `getContentAsText({ annotated })` option that inlines tracked changes and comment anchors as `<ins>`/`<del>`/`<comment>` tags, and a read-only `getNotesAsText` surfacing header/footer and footnote/endnote text. The React and Vue `DocxEditorRef` gain a matching read surface — `getTrackedChanges`, `getCommentAnchors`, `getSelectionText`, and `getPageText(page)` — so a live-editor agent tool can read the current document state without a fresh AI-edit snapshot. `@stll/folio-agents` now exports `parseSuggestChangesInput` / `parseAddCommentInput` — the `suggest_changes` / `add_comment` argument-validation rules factored out of `executeFolioToolCall`, for hosts with their own review-queue UX — and its live-editor bridge now reads real tracked changes and comment anchors from that new `DocxEditorRef` surface instead of reporting them as unavailable.

### Patch Changes

- Updated dependencies [[`ca9d64b`](https://github.com/stella/folio/commit/ca9d64bd36bbce78b3dad9aab72092cceebc4919), [`439a8ea`](https://github.com/stella/folio/commit/439a8ea12227aafc49e99693755fc03577d4b54c), [`ce39713`](https://github.com/stella/folio/commit/ce397130771ec8f1271b7d0cfde9e6dc0367e857)]:
  - @stll/folio-core@0.3.0

## 0.3.0

### Minor Changes

- [#54](https://github.com/stella/folio/pull/54) [`444c11b`](https://github.com/stella/folio/commit/444c11b6c165864bbea59e4ed54c498d1b6fa02e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Readable template directive rails: gutter rails indent by nesting depth
  within a margin-confined budget, colour by directive kind (loop violet,
  condition teal, theme-aware), render quiet by default and emphasize the
  block containing the caret or under the pointer, split into per-page
  segments so no rail crosses a page header, footer, or gap, and pair
  opener/closer chips on hover with a hint of what a bare closer closes.
  Adds pure helpers `computeBlockDepths` (core) and gutter geometry
  measurement to the paged-layout range projection.

### Patch Changes

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

- Updated dependencies [[`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8), [`444c11b`](https://github.com/stella/folio/commit/444c11b6c165864bbea59e4ed54c498d1b6fa02e), [`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8), [`e44ba30`](https://github.com/stella/folio/commit/e44ba300c0978b9e2836bafcfe84cd06491e87bf), [`c40cb00`](https://github.com/stella/folio/commit/c40cb00390db3abf109c3764645a87451fb6a249)]:
  - @stll/folio-core@0.2.0

## 0.2.0

### Minor Changes

- Add controlled comment props for collaboration sync, `scrollToParaId`,
  `renderAsync`, and Hebrew, Hindi, Turkish, and Simplified Chinese UI locales.

### Patch Changes

- Updated dependencies []:
  - @stll/folio-core@0.1.3
