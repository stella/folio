# @stll/folio-react

## 0.7.0

### Minor Changes

- [#138](https://github.com/stella/folio/pull/138) [`502f114`](https://github.com/stella/folio/commit/502f1140ba3fedb6b15544c85c1244ca50d0c28c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Port the Insert Symbol dialog to the React adapter, reaching full parity with Vue. The symbol catalog and search now live in a shared framework-neutral `@stll/folio-core/symbols` module (six categories, `filterSymbols`); the Vue dialog is refactored onto it so both adapters share one source of truth instead of duplicating the catalog. React gains `InsertSymbolDialog` (exported from the package root), an "Insert Symbol" toolbar button (`onInsertSymbol`), and inserts the chosen character at the cursor.

- [#104](https://github.com/stella/folio/pull/104) [`6e95dc2`](https://github.com/stella/folio/commit/6e95dc2b9f2bc8060d0c3933fdebab7bbf7e0f9d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Adopt versioned document operation batches across agent and live editor bridges.

### Patch Changes

- [#132](https://github.com/stella/folio/pull/132) [`6d136cf`](https://github.com/stella/folio/commit/6d136cf795b9f088e726e0524f936b3a5135db85) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add dry-run previews for document operation batches.

- [#166](https://github.com/stella/folio/pull/166) [`2234792`](https://github.com/stella/folio/commit/223479229acfa7f51185cf18408528a0e3df9790) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add typed affected-target receipts to document operation results.

- [#152](https://github.com/stella/folio/pull/152) [`95aca77`](https://github.com/stella/folio/commit/95aca7790f864974de8857b23e25ffcf732ead89) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add structured diagnostics to document operation results.

- [#143](https://github.com/stella/folio/pull/143) [`8f0701e`](https://github.com/stella/folio/commit/8f0701ed1f2b527ba21820f71a82c623b473d322) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Hoist the `getPageTextFromLayout` and AI-edit block-range helpers (`resolveFolioAIBlockRange`, `clampRangeToDocSize`) into framework-neutral core modules (`@stll/folio-core/paged-layout/pageText`, `@stll/folio-core/ai-edits/blockRange`). Both were previously duplicated verbatim in the React and Vue adapters; they now share one implementation (and one test suite) in core, so the two adapters can never drift on page-text extraction or block-range resolution.

- [#124](https://github.com/stella/folio/pull/124) [`411f5e4`](https://github.com/stella/folio/commit/411f5e433c69cb04bc3c2ea4b8d141faa791d80b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce typing latency by coalescing incremental layout on the next animation frame.

- [#170](https://github.com/stella/folio/pull/170) [`8423a10`](https://github.com/stella/folio/commit/8423a1090795bac001c7d0903d0d3ba8ec31b140) Thanks [@cursor](https://github.com/apps/cursor)! - Compile React components with React Compiler.

- [#164](https://github.com/stella/folio/pull/164) [`f6d6a7c`](https://github.com/stella/folio/commit/f6d6a7caab18a10b354e60dcae845d77f3a7d8a0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor mirrored left and right margins on facing document pages.

- [#116](https://github.com/stella/folio/pull/116) [`476fa19`](https://github.com/stella/folio/commit/476fa1918f041d3c9ffcf097b0609d1b0d37b069) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve Aptos document layout and preserve document-default spacing on empty paragraphs.

- [#144](https://github.com/stella/folio/pull/144) [`be03ce5`](https://github.com/stella/folio/commit/be03ce5245ca8829fbbd75b23865a3112827c0c6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Translate the display-mode select trigger label and stabilize object, array, function, and JSX props across React component boundaries to preserve memoization in the published build.

- [#126](https://github.com/stella/folio/pull/126) [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Unify the reachable zoom range across adapters. A new `@stll/folio-core/utils/zoom` module exports the canonical `ZOOM_MIN` (0.25), `ZOOM_MAX` (4), and `ZOOM_STEP` (0.1); React and Vue now source their clamp from it. The Vue adapter previously clamped zoom to 0.5-2x and now matches React at 0.25-4x. Each adapter's curated toolbar zoom-level dropdown (50-200%) is unchanged and remains an intentional subset of the reachable range.

- [#130](https://github.com/stella/folio/pull/130) [`317fa96`](https://github.com/stella/folio/commit/317fa96ab0b3254b9780ca621935b6c38787a204) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add optional atomic document operation batches.

- Updated dependencies [[`6d136cf`](https://github.com/stella/folio/commit/6d136cf795b9f088e726e0524f936b3a5135db85), [`33d2d04`](https://github.com/stella/folio/commit/33d2d04797f26c5391753feaddd942a84db53f96), [`b90687b`](https://github.com/stella/folio/commit/b90687b2bf6907de8cef2c7156d8691db7df45d8), [`2234792`](https://github.com/stella/folio/commit/223479229acfa7f51185cf18408528a0e3df9790), [`30b086e`](https://github.com/stella/folio/commit/30b086eb4e9779767fc94828633a0a15de8de0e4), [`7cfef90`](https://github.com/stella/folio/commit/7cfef90a1e3bc4d7324d13025eeb72bfccf7f4ee), [`50990a0`](https://github.com/stella/folio/commit/50990a0da74086f8c283bc99a68cc7db2bfaed82), [`33d2d04`](https://github.com/stella/folio/commit/33d2d04797f26c5391753feaddd942a84db53f96), [`33538c6`](https://github.com/stella/folio/commit/33538c633d1bd206996260de8f618ed40629f2c5), [`2eb96d5`](https://github.com/stella/folio/commit/2eb96d5d8d0bf3c7cfa933201c6310184f673825), [`7bd5dd8`](https://github.com/stella/folio/commit/7bd5dd8a623d5dfae28c5b43e507c4ab4c005507), [`d4a40b2`](https://github.com/stella/folio/commit/d4a40b2e28ed37390dc0a8ffbc759ba1d068eb4f), [`bf8e841`](https://github.com/stella/folio/commit/bf8e8411e6e812f7dba5ccb998960e869607507e), [`63eee6c`](https://github.com/stella/folio/commit/63eee6ca8b0188283255c8a6cb69551016b2d5dc), [`823b172`](https://github.com/stella/folio/commit/823b172cfa3fcb3b7339fcf60b6705db7b3662cb), [`58e66c3`](https://github.com/stella/folio/commit/58e66c3ec5e56b7809c8694249c6c69509e9830a), [`c864d1c`](https://github.com/stella/folio/commit/c864d1c0071cb5c7fec1fd59b68aba22def02083), [`95aca77`](https://github.com/stella/folio/commit/95aca7790f864974de8857b23e25ffcf732ead89), [`7fcb7d7`](https://github.com/stella/folio/commit/7fcb7d7242edbf61dac44aefcad3aa6990a13b71), [`8f0701e`](https://github.com/stella/folio/commit/8f0701ed1f2b527ba21820f71a82c623b473d322), [`411f5e4`](https://github.com/stella/folio/commit/411f5e433c69cb04bc3c2ea4b8d141faa791d80b), [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736), [`7bdd026`](https://github.com/stella/folio/commit/7bdd026483d29adcc1b57d5916d4e850c2e270d5), [`50b61d1`](https://github.com/stella/folio/commit/50b61d195834f00c1f07c27b3867a78803e02a91), [`f0d1cba`](https://github.com/stella/folio/commit/f0d1cba5ca7a1830aae7703e226e731fae76524b), [`750a137`](https://github.com/stella/folio/commit/750a1378c644314b7e478b593b76d41a1dbdb4cf), [`0f6f547`](https://github.com/stella/folio/commit/0f6f5472c858a1e67135ed9205b1600b62a61314), [`f7de365`](https://github.com/stella/folio/commit/f7de36526d276f4c2d5e684da3a468990c075ac4), [`f6d6a7c`](https://github.com/stella/folio/commit/f6d6a7caab18a10b354e60dcae845d77f3a7d8a0), [`3958226`](https://github.com/stella/folio/commit/3958226cf7047b9b16d506f4e18548ecf23950e3), [`1a278bf`](https://github.com/stella/folio/commit/1a278bfe33aeab6f9c4b92923223712d82617547), [`1742abe`](https://github.com/stella/folio/commit/1742abeb2f00f3215767bc3898110463844e279a), [`685ee93`](https://github.com/stella/folio/commit/685ee93fac2b5cb573c9569cd85b9c98fe3c9bca), [`8d9808c`](https://github.com/stella/folio/commit/8d9808c9fd621011b4f73a3bf9ac6e7ee3391a80), [`6997eee`](https://github.com/stella/folio/commit/6997eee34255bfa28ef3048e7c8bcb443c0ac091), [`f1cf28b`](https://github.com/stella/folio/commit/f1cf28bccb5fd5742e45ccdcd651b500c1234950), [`3130e5f`](https://github.com/stella/folio/commit/3130e5f616254e6bb127e4c410871c05f6bb386f), [`2f1a8cb`](https://github.com/stella/folio/commit/2f1a8cbf922b6ddf4fd93f105bbdc95d3ac0faa4), [`8a3888c`](https://github.com/stella/folio/commit/8a3888c71c6f68970342a6c6cde9af7baf4c1565), [`d736b57`](https://github.com/stella/folio/commit/d736b57f07953e8faf8793d3324912fa2b8e52f0), [`476fa19`](https://github.com/stella/folio/commit/476fa1918f041d3c9ffcf097b0609d1b0d37b069), [`1fc3f2f`](https://github.com/stella/folio/commit/1fc3f2f748f4330838a56af20f7a3dbf5e5d8959), [`1f74701`](https://github.com/stella/folio/commit/1f74701d3fc2e788a078f8620f006c294c6b8b4e), [`9697a47`](https://github.com/stella/folio/commit/9697a47eb1c5b93acb374e3ed53e1bdbda47b65e), [`a5d496f`](https://github.com/stella/folio/commit/a5d496f460c56eec6c7d5c3dd45f528f13f3c144), [`502f114`](https://github.com/stella/folio/commit/502f1140ba3fedb6b15544c85c1244ca50d0c28c), [`c01b5b0`](https://github.com/stella/folio/commit/c01b5b0f70ad38c9ff878f8396a418de563ff06a), [`41bc093`](https://github.com/stella/folio/commit/41bc093036e312971532f481d4320e3c63c57f51), [`174d0ce`](https://github.com/stella/folio/commit/174d0ce25de076e9a51963547bb37b059b5b21d3), [`f03fa89`](https://github.com/stella/folio/commit/f03fa8966bcee91ed03e33a1bc9f17513960b471), [`0caa2e6`](https://github.com/stella/folio/commit/0caa2e680e0fe125811215ad7ac4a38194ffa673), [`ffc12e6`](https://github.com/stella/folio/commit/ffc12e6142b03abd4ce27983019b9256719c7b2b), [`9af69ca`](https://github.com/stella/folio/commit/9af69ca2d63fab2f6795b388184f031b61a87545), [`dc188a4`](https://github.com/stella/folio/commit/dc188a42f21db9cccdd92f292a979c71d8be358b), [`4377bf0`](https://github.com/stella/folio/commit/4377bf0db4e7390c409a232852b4d20a533edc08), [`4de7540`](https://github.com/stella/folio/commit/4de75406dea5e63b96cbf1620429b198ec6f9a33), [`c7690f9`](https://github.com/stella/folio/commit/c7690f9e77b1a048d8ea6c77676436e3cd191bad), [`0ee5fba`](https://github.com/stella/folio/commit/0ee5fba1231d11f638b56c2b78c63c8479a1fed5), [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736), [`317fa96`](https://github.com/stella/folio/commit/317fa96ab0b3254b9780ca621935b6c38787a204), [`1bea7a9`](https://github.com/stella/folio/commit/1bea7a9a4a4ebf8ff263735460a27c2d3ec1241b), [`9af69ca`](https://github.com/stella/folio/commit/9af69ca2d63fab2f6795b388184f031b61a87545)]:
  - @stll/folio-core@0.5.0

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
