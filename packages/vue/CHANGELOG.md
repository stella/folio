# @stll/folio-vue

## 0.4.0

### Minor Changes

- [#91](https://github.com/stella/folio/pull/91) [`521879e`](https://github.com/stella/folio/commit/521879e5b3db5c91e8a68dec0de14e31cc964557) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add Vue dialog parity for watermark, insert image, paste special, and split cell workflows, and auto-register Vue dialogs in Nuxt.

### Patch Changes

- Updated dependencies [[`989999c`](https://github.com/stella/folio/commit/989999cc3629c434a36b16bffee8cba8eb2171b4), [`63731b6`](https://github.com/stella/folio/commit/63731b61e9be9601adf52a6faceaa7cd1ee9fbc4), [`521879e`](https://github.com/stella/folio/commit/521879e5b3db5c91e8a68dec0de14e31cc964557), [`b637aa4`](https://github.com/stella/folio/commit/b637aa44b5f6705affa859c037af96b38df4360d)]:
  - @stll/folio-core@0.4.0

## 0.3.0

### Minor Changes

- [#81](https://github.com/stella/folio/pull/81) [`439a8ea`](https://github.com/stella/folio/commit/439a8ea12227aafc49e99693755fc03577d4b54c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Broaden the AI-edits read surface for live-editor reuse. Core gains pure `getTrackedChangesFromDoc` / `getCommentAnchorsFromDoc` readers and a `buildAnnotatedBlockText` redline renderer over any ProseMirror `doc`, a `getContentAsText({ annotated })` option that inlines tracked changes and comment anchors as `<ins>`/`<del>`/`<comment>` tags, and a read-only `getNotesAsText` surfacing header/footer and footnote/endnote text. The React and Vue `DocxEditorRef` gain a matching read surface — `getTrackedChanges`, `getCommentAnchors`, `getSelectionText`, and `getPageText(page)` — so a live-editor agent tool can read the current document state without a fresh AI-edit snapshot. `@stll/folio-agents` now exports `parseSuggestChangesInput` / `parseAddCommentInput` — the `suggest_changes` / `add_comment` argument-validation rules factored out of `executeFolioToolCall`, for hosts with their own review-queue UX — and its live-editor bridge now reads real tracked changes and comment anchors from that new `DocxEditorRef` surface instead of reporting them as unavailable.

- [#68](https://github.com/stella/folio/pull/68) [`9f00464`](https://github.com/stella/folio/commit/9f004647fab87faab214a7fbdd6a1543810f2e79) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire the tractable remaining chrome/callback props on the Vue `DocxEditor`,
  moving them from deferred to paired with the React adapter: `showRuler` +
  `rulerUnit` (renders the horizontal and vertical rulers, gated on `!readOnly`,
  with live margin / indent / tab-stop editing), `autoOpenReviewSidebar`,
  `initialScrollTop` + `onScrollTopChange`, `onFontsLoaded`,
  `onCompatibilityChange`, `onSelectionTextChange`, `placeholder` /
  `loadingIndicator` (nodes or `#placeholder` / `#loading-indicator` slots), and
  `customContextMenuItems` + `onCustomContextAction` (host-injected right-click
  entries that lead the menu, with the selection range captured at open time).
  Genuinely-hard or React-inert props (header/footer editing, template overlay,
  anonymization overlays, host FontFace registration, keep-previous-document-
  while-loading, margin guides, and the clipboard / review-control / print-button
  flags React itself leaves unwired) stay deferred with updated reasons.

- [#66](https://github.com/stella/folio/pull/66) [`99c96c2`](https://github.com/stella/folio/commit/99c96c25d11f3f9eaa23c0906efcf5919ad2a8d9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire the comment + tracked-change lifecycle into `DocxEditor.vue`. The sidebar
  now populates from the document: comments are seeded from
  `document.package.document.comments` on load, and tracked-change cards derive
  from core's `extractTrackedChanges(editorState)` (the Vue `TrackedChangeEntry`
  now re-exports core's type, removing the optional-field mismatch that left
  cards empty). Two new composables built on core's comment ops drive mutations:
  `useCommentManagement` (add reply, resolve/unresolve, delete, tracked-change
  reply, and accept/reject by revision id via `findChangeRange` + range-based
  `acceptChange`/`rejectChange`) and `useCommentLifecycle` (floating add-comment
  button, pending-highlight range, submit/cancel). Every mutation fires
  `onCommentsChange` (and a `comments-change` emit). `comments` and
  `onCommentsChange` move from `deferredInVue` to `paired`.

- [#69](https://github.com/stella/folio/pull/69) [`b5013db`](https://github.com/stella/folio/commit/b5013db726f806cdd0df4eae81dfac974b756e70) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Port the anonymization + template-directive decoration overlays to the Vue
  adapter. Both project the existing core plugins' decoration ranges onto
  page-relative rects via `projectRangesToRects` (DOM-rect primary path, layout
  fallback) and paint them over the pages. `AnonymizationRectsOverlay.vue`
  highlights redaction matches, forwards clicks, and scrolls the selected term
  into view; `TemplateDirectivesOverlay.vue` tints `{{…}}` markers and draws the
  per-page block gutter rails. The anonymization, template-directive, slash-menu,
  and template-preview plugins are now registered in the hidden-editor pipeline
  (template plugins gated on `showTemplateDirectives`, reconfigured live on
  toggle). Moves `onAnonymizationMatchesChange`, `onAnonymizationTermClick`,
  `selectedAnonymizationCanonical`, `anonymizationSelectionSeq`,
  `showTemplateDirectives`, `onSlashMenuChange`, and `onSlashMenuKeyAction` from
  deferred to paired in the cross-adapter parity contract.

- [#70](https://github.com/stella/folio/pull/70) [`2505ac6`](https://github.com/stella/folio/commit/2505ac6682aafb99c00ea079accfd6001550d95a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add standalone `FormattingBar` and `ZoomControl` Vue components, matching the
  React adapter's public surface. Both are controlled and own no editor view, so a
  host can render them outside `<DocxEditor>`. `FormattingBar` (`FormattingBarProps`)
  composes the existing Vue pickers into the minimal rail (undo/redo | style |
  B/I/U | color | align/lists/indent) and emits `format` / `undo` / `redo`;
  `ZoomControl` (`ZoomControlProps`, `ZoomLevel`) is a compact zoom-level dropdown.
  This closes the corresponding React/Vue export-parity divergence entries.

- [#75](https://github.com/stella/folio/pull/75) [`d2a3408`](https://github.com/stella/folio/commit/d2a34089a0e019ebf3df67afc8fbb40489d117db) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Implement `DocxEditorRef.hasPendingChanges()` in the Vue adapter (previously
  stubbed to `false`). It now returns whether the live editor has edits not yet
  serialized by `save()`, mirroring the React adapter's set/clear points: a
  doc-dirty flag set on every doc-changing transaction and cleared on a successful
  `save()` and on document load/swap, OR-ed with a comment-list dirty flag set on
  comment mutations (add/reply/resolve/unresolve/delete) and cleared on load and
  save. Moves `hasPendingChanges` from deferred to paired in the cross-adapter
  parity contract.

- [#64](https://github.com/stella/folio/pull/64) [`e38aa31`](https://github.com/stella/folio/commit/e38aa31f69207cae8f382a91a6452e444304c6f1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire the Vue title-bar MenuBar to the real `MenuDropdown` / `TableGridInline`
  primitives (replacing the inline render-null stubs) and route the document-insert
  flow end to end. `MenuBar` items now drive image / table / page-break /
  table-of-contents insertion through `DocxEditor.vue`, which prefers the host
  `onInsertImage` / `onInsertTable` / `onInsertPageBreak` / `onInsertTOC` props when
  provided and otherwise falls back to the core view-level helpers
  (`insertImageFromFile`, `insertTableInView`, `insertPageBreakInView`,
  `insertTableOfContentsInView`). `showTableInsert` now gates the Insert > Table
  menu item, mirroring the React toolbar.

- [#78](https://github.com/stella/folio/pull/78) [`ea1ff9d`](https://github.com/stella/folio/commit/ea1ff9d0fb8d93a5fc43d1bc9a353b6955207283) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire `DocxEditorRef.getEditorRef()` on the Vue adapter, moving it from
  deferred to paired with React's `PagedEditorRef`. The Vue package has no
  ported `PagedEditor` sub-component to source a handle from, so
  `useDocxEditorRefApi.ts` synthesizes an equivalent `PagedEditorRef`-shaped
  object from primitives it already holds: `getEditor` / `getDocument` /
  `getState` / `getView` / `ensureView` / `focus` / `blur` / `isFocused` /
  `dispatch` / `undo` / `redo` / `canUndo` / `canRedo` / `setSelection` /
  `getLayout` / `relayout` delegate to the headless `FolioEditor` controller;
  `scrollToPosition` reuses the pages-scroll helper; `scrollToPage` /
  `scrollToParaId` reuse the existing top-level ref implementations;
  `getPageNumberForPmPos` resolves through the same layout helper
  `getCurrentPage` already uses. `getHfView` stays a documented no-op (returns
  `null`) since Vue has no persistent hidden header/footer `EditorView` yet.

- [#72](https://github.com/stella/folio/pull/72) [`34c4f0b`](https://github.com/stella/folio/commit/34c4f0bfa460d8c5757008ca2434972f089e1fb8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Squeeze the Vue adapter's remaining prop deferrals to 58/64 paired. Wire three
  real props: `toolbarExtra` (the `VNodeChild` prop now renders into the toolbar's
  `#toolbar-extra` outlet, with a host-provided slot taking precedence),
  `preserveDocumentWhileLoading` (suppresses the loading interstitial on a document
  swap once a document has painted, so the previous pages stay visible until the
  new layout replaces them), and `fonts` (ports React's host `FontFace`
  registration + font-ready re-layout via a DOM-only `utils/hostFonts` module).

  Reclassify six no-op-in-both props to paired after verifying React declares but
  never consumes them: `showMarginGuides`, `marginGuideColor`, `showPrintButton`,
  `onCopy`, `onCut`, `onPaste`. `showReviewControls` stays deferred (a real gap:
  React renders review controls the Vue toolbar lacks).

- [#67](https://github.com/stella/folio/pull/67) [`b2ecfa9`](https://github.com/stella/folio/commit/b2ecfa973ebf45cfed426501fd57abb0d64cb439) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire the AI-edit and content-control `DocxEditorRef` methods in the Vue adapter
  over `@stll/folio-core`, mirroring the React adapter. `createAIEditSnapshot`,
  `applyAIEditOperations`, `acceptAIEditOperation`, `rejectAIEditOperation`,
  `scrollToAIEditOperation`, `scrollToBlock`, `getContentControls`,
  `scrollToContentControl`, `setContentControlContent`, `setContentControlValue`,
  and `removeContentControl` were previously stubbed to `null` / `false` / `[]`;
  they now drive the live editor view. Only `getEditorRef` and `hasPendingChanges`
  remain deferred pending the Vue `PagedEditor` component and PM-serialized-vs-live
  tracking.

- [#74](https://github.com/stella/folio/pull/74) [`f36f171`](https://github.com/stella/folio/commit/f36f171def307d41c5386291f82e7f4889a348ac) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire the review controls in the Vue toolbar, gated on `showReviewControls`
  (default true), matching the React adapter. Adds a track-changes toggle
  (flips editing<->suggesting through the existing `setSuggestionMode` path) and a
  markup display-mode selector (All Markup / Simple / No Markup / Original) that
  applies the same `folio-root--<mode>` root class and tracked-change display CSS
  as React. `showReviewControls` moves from deferred to paired in the
  cross-adapter parity contract.

- [#73](https://github.com/stella/folio/pull/73) [`94da1d4`](https://github.com/stella/folio/commit/94da1d404ac41c03161a7c9209182d5acab16d9e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire selective-save feature flags and the tripwire callback into the Vue
  `DocxEditor` save path, matching the React adapter. `featureFlags` is threaded
  into `useDocxEditor`'s `save()`: when `selectiveSave` is on it runs
  `attemptSelectiveSave` (patching only changed paragraphs, with the shared
  `ParagraphChangeTracker` inputs) and falls back to a full repack; the ref
  `save({ selective })` option now gates that path. When `selectiveSaveTripwire`
  is on it also computes the full repack, runs `compareSelectiveVsFull`, and fires
  `onSelectiveSaveTripwire(result)` for observability. The comparison never blocks
  or poisons the save. Both props move from deferred to paired in the cross-adapter
  parity contract.

- [#77](https://github.com/stella/folio/pull/77) [`5f81c95`](https://github.com/stella/folio/commit/5f81c95a1673f0884e1078ec1bf99283680a5318) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Close the Vue `components` UI-injection gap so a host can override all ten of
  folio's chrome primitives, matching React: `FolioUIComponents` (and
  `DEFAULT_COMPONENTS`) grow from `Button`/`ColorPicker`/`Popover`/`Menu` to the
  full set, adding `Dialog`, `Select`, `Input`, `Checkbox`, `DatePickerPopover`,
  and `OutlineRail` with a real Vue default per primitive (each collapsed into
  one monolithic, data-driven component rather than React's base-ui part-object
  shape — the pattern the existing four already used). Chrome consumers now
  resolve their primitive through `useFolioUI()` instead of a static import:
  `MenuBar` (Menu), 8 previously-static `Popover` consumers
  (`EditingModeDropdown`, `ReviewControls`, `TableGridPicker`,
  `TableBorderWidthPicker`, `IconGridDropdown`, `TableBorderPicker`,
  `TableMoreDropdown`, `AlignmentButtons`), `ZoomControl` / `StylePicker`
  (Select), `FindReplaceDialog` (Button/Input/Checkbox), the five modal dialogs
  — `PageSetupDialog`, `ImagePropertiesDialog`, `FootnotePropertiesDialog`,
  `ImagePositionDialog`, `TablePropertiesDialog` (Dialog), and
  `DocumentOutline`'s item list (OutlineRail). `components` moves from
  `deferredInVue` to `paired` in the parity contract. New default components and
  the `Folio*Props` contract types are exported from the `@stll/folio-vue/ui`
  subpath. Minor fix bundled in: Escape now closes all five modal dialogs
  (`ImagePropertiesDialog` previously swallowed the keydown before any handler
  saw it).

- [#71](https://github.com/stella/folio/pull/71) [`b249376`](https://github.com/stella/folio/commit/b2493769f956d5f4901f6a297d4d4678ac97ba7d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add the Vue UI-injection layer (`FolioUIProvider` / `provideFolioUI` / `useFolioUI`
  / `resolveFolioComponents` / `DEFAULT_COMPONENTS`), mirroring the React adapter's
  `FolioUIComponents` contract. `DocxEditor` now provides the resolved primitive map
  from its `components` prop, and the toolbar / formatting-bar chrome resolve their
  `ColorPicker` through the provider, so a host override takes effect. Exports the
  `FolioUIComponents`, `ColorPreset`, `FolioButtonProps`, and `OutlineItem` types plus
  the `FolioUIProvider` component, closing the corresponding React↔Vue export-parity
  gaps.

- [#63](https://github.com/stella/folio/pull/63) [`d1ac0db`](https://github.com/stella/folio/commit/d1ac0dbb16ccd429a693f7f8ab0754a87d059bf5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire ported composables and real pickers into `DocxEditor.vue`, replacing the
  shell-time PORT-BLOCKED stubs. Pages-area pointer gestures (`usePagesPointer`:
  multi-click, drag-select, table quick-insert button, hyperlink popup, page
  indicator), image actions + selection overlay (`useImageActions` +
  `ImageSelectionOverlay`), right-click menus (`useContextMenus`), hyperlink
  management (`useHyperlinkManagement`), the table context toolbar (`TableToolbar`
  with its border/fill/more pickers), and the comment sidebar (`UnifiedSidebar`
  now renders controlled comments via `useCommentSidebarItems`) are functional.
  The toolbar image group (`ImageWrapDropdown`/`ImageTransformDropdown`) and the
  Insert Table dialog's style gallery (`TableStyleGallery`) are un-stubbed. 2
  props (`enableWheelZoom`, `onSave`) and 1 ref member (`scrollToParaId`) move
  from `deferredInVue` to `paired`.

### Patch Changes

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

- [#79](https://github.com/stella/folio/pull/79) [`a45911c`](https://github.com/stella/folio/commit/a45911c27b527c09d6a1bdfa376e8dcad2a2e08d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fix the Document Outline panel always showing "No headings found": `outlineHeadings`
  was declared but never populated. Wire the existing `useOutlineSidebar` composable
  into `DocxEditor.vue` so opening the outline (via the toggle button, toolbar, or File
  menu) collects headings from the document, re-collects them on further edits while the
  panel stays open, and clicking a heading scrolls the visible pages to it.

  Fix Find/Replace having no way to open: `useKeyboardShortcuts` existed but was never
  invoked from `DocxEditor.vue`, so Ctrl+F / Ctrl+H / Ctrl+K never opened their dialogs.
  Invoke the composable from the editor shell, and add a Find & Replace entry to the
  File menu as a discoverable, mouse-driven entry point alongside the keyboard shortcut.

- [#61](https://github.com/stella/folio/pull/61) [`3900c6b`](https://github.com/stella/folio/commit/3900c6b83e7062d3204922ba2d88adb0c474f064) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fix the published subpath exports so consumers resolve real artifacts: the
  `./composables`, `./dialogs`, and `./styles` `types` conditions pointed at
  `dist/*.d.ts` files the build never emits (the declarations flatten to
  `dist/composables/index.d.ts`, `dist/components/dialogs/index.d.ts`, and
  `dist/styles/index.d.ts`), and `./messages` pointed at a `./src/*.ts` source
  file absent from the published tarball. `./messages` is now a built entry
  (`dist/messages.js` / `.cjs` + `dist/i18n/messages.d.ts`).
- Updated dependencies [[`ca9d64b`](https://github.com/stella/folio/commit/ca9d64bd36bbce78b3dad9aab72092cceebc4919), [`439a8ea`](https://github.com/stella/folio/commit/439a8ea12227aafc49e99693755fc03577d4b54c), [`ce39713`](https://github.com/stella/folio/commit/ce397130771ec8f1271b7d0cfde9e6dc0367e857)]:
  - @stll/folio-core@0.3.0

## 0.2.0

### Minor Changes

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

- Updated dependencies [[`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8), [`444c11b`](https://github.com/stella/folio/commit/444c11b6c165864bbea59e4ed54c498d1b6fa02e), [`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8), [`e44ba30`](https://github.com/stella/folio/commit/e44ba300c0978b9e2836bafcfe84cd06491e87bf), [`c40cb00`](https://github.com/stella/folio/commit/c40cb00390db3abf109c3764645a87451fb6a249)]:
  - @stll/folio-core@0.2.0
