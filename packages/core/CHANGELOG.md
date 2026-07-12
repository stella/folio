# @stll/folio-core

## 0.5.0

### Minor Changes

- [#132](https://github.com/stella/folio/pull/132) [`6d136cf`](https://github.com/stella/folio/commit/6d136cf795b9f088e726e0524f936b3a5135db85) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add dry-run previews for document operation batches.

- [#140](https://github.com/stella/folio/pull/140) [`b90687b`](https://github.com/stella/folio/commit/b90687b2bf6907de8cef2c7156d8691db7df45d8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add range comments and constrained inline formatting operations.

- [#166](https://github.com/stella/folio/pull/166) [`2234792`](https://github.com/stella/folio/commit/223479229acfa7f51185cf18408528a0e3df9790) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add typed affected-target receipts to document operation results.

- [#107](https://github.com/stella/folio/pull/107) [`7cfef90`](https://github.com/stella/folio/commit/7cfef90a1e3bc4d7324d13025eeb72bfccf7f4ee) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Report supported mutation modes for each document operation type.

- [#123](https://github.com/stella/folio/pull/123) [`2eb96d5`](https://github.com/stella/folio/commit/2eb96d5d8d0bf3c7cfa933201c6310184f673825) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add serialized block text preconditions to document operations.

- [#152](https://github.com/stella/folio/pull/152) [`95aca77`](https://github.com/stella/folio/commit/95aca7790f864974de8857b23e25ffcf732ead89) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add structured diagnostics to document operation results.

- [#143](https://github.com/stella/folio/pull/143) [`8f0701e`](https://github.com/stella/folio/commit/8f0701ed1f2b527ba21820f71a82c623b473d322) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Hoist the `getPageTextFromLayout` and AI-edit block-range helpers (`resolveFolioAIBlockRange`, `clampRangeToDocSize`) into framework-neutral core modules (`@stll/folio-core/paged-layout/pageText`, `@stll/folio-core/ai-edits/blockRange`). Both were previously duplicated verbatim in the React and Vue adapters; they now share one implementation (and one test suite) in core, so the two adapters can never drift on page-text extraction or block-range resolution.

- [#136](https://github.com/stella/folio/pull/136) [`750a137`](https://github.com/stella/folio/commit/750a1378c644314b7e478b593b76d41a1dbdb4cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add stable text-range handles and exact range replacement operations.

- [#106](https://github.com/stella/folio/pull/106) [`0f6f547`](https://github.com/stella/folio/commit/0f6f5472c858a1e67135ed9205b1600b62a61314) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Validate serialized document operation batches before execution.

- [#100](https://github.com/stella/folio/pull/100) [`685ee93`](https://github.com/stella/folio/commit/685ee93fac2b5cb573c9569cd85b9c98fe3c9bca) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add versioned document operation batches and capability discovery.

- [#138](https://github.com/stella/folio/pull/138) [`502f114`](https://github.com/stella/folio/commit/502f1140ba3fedb6b15544c85c1244ca50d0c28c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Port the Insert Symbol dialog to the React adapter, reaching full parity with Vue. The symbol catalog and search now live in a shared framework-neutral `@stll/folio-core/symbols` module (six categories, `filterSymbols`); the Vue dialog is refactored onto it so both adapters share one source of truth instead of duplicating the catalog. React gains `InsertSymbolDialog` (exported from the package root), an "Insert Symbol" toolbar button (`onInsertSymbol`), and inserts the chosen character at the cursor.

- [#174](https://github.com/stella/folio/pull/174) [`9af69ca`](https://github.com/stella/folio/commit/9af69ca2d63fab2f6795b388184f031b61a87545) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Version comparison upgrades and a redline generator. `compareDocxVersions` now detects relocated blocks (`movedFrom`/`movedTo` pairs sharing a `moveGroupId`) instead of reporting them as unrelated delete + insert, and reports text-equal blocks whose run formatting differs as `formatChanged` with the changed property names. New `generateRedlineDocx(base, revised)` produces a third `.docx` whose base → revised differences are recorded as real Word tracked changes (`w:ins`/`w:del`), reusing the comparer alignment and the headless tracked-changes apply path. `formatVersionDiffForLLM` renders the new change types.

- [#109](https://github.com/stella/folio/pull/109) [`dc188a4`](https://github.com/stella/folio/commit/dc188a42f21db9cccdd92f292a979c71d8be358b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Report unsupported document operation modes separately from unsupported blocks.

- [#156](https://github.com/stella/folio/pull/156) [`4de7540`](https://github.com/stella/folio/commit/4de75406dea5e63b96cbf1620429b198ec6f9a33) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose document version comparison from folio-core while preserving the agents API.

- [#141](https://github.com/stella/folio/pull/141) [`0ee5fba`](https://github.com/stella/folio/commit/0ee5fba1231d11f638b56c2b78c63c8479a1fed5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add typed discovery and reads for document stories.

- [#126](https://github.com/stella/folio/pull/126) [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Unify the reachable zoom range across adapters. A new `@stll/folio-core/utils/zoom` module exports the canonical `ZOOM_MIN` (0.25), `ZOOM_MAX` (4), and `ZOOM_STEP` (0.1); React and Vue now source their clamp from it. The Vue adapter previously clamped zoom to 0.5-2x and now matches React at 0.25-4x. Each adapter's curated toolbar zoom-level dropdown (50-200%) is unchanged and remains an intentional subset of the reachable range.

- [#130](https://github.com/stella/folio/pull/130) [`317fa96`](https://github.com/stella/folio/commit/317fa96ab0b3254b9780ca621935b6c38787a204) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add optional atomic document operation batches.

### Patch Changes

- [#139](https://github.com/stella/folio/pull/139) [`33d2d04`](https://github.com/stella/folio/commit/33d2d04797f26c5391753feaddd942a84db53f96) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor authored document default tab intervals in paragraph layout and painting.

- [#160](https://github.com/stella/folio/pull/160) [`30b086e`](https://github.com/stella/folio/commit/30b086eb4e9779767fc94828633a0a15de8de0e4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fix justified spacing on hanging first lines that contain tabs.

- [#134](https://github.com/stella/folio/pull/134) [`50990a0`](https://github.com/stella/folio/commit/50990a0da74086f8c283bc99a68cc7db2bfaed82) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep authored body margins independent from header/footer paint bounds and carry keep-with-next chains across empty separators.

- [#139](https://github.com/stella/folio/pull/139) [`33d2d04`](https://github.com/stella/folio/commit/33d2d04797f26c5391753feaddd942a84db53f96) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve standard list continuation wrapping while matching custom hanging indents.

- [#151](https://github.com/stella/folio/pull/151) [`33538c6`](https://github.com/stella/folio/commit/33538c633d1bd206996260de8f618ed40629f2c5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce DOCX parsing time by skipping unused XML callback paths.

- [#117](https://github.com/stella/folio/pull/117) [`7bd5dd8`](https://github.com/stella/folio/commit/7bd5dd8a623d5dfae28c5b43e507c4ab4c005507) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor authored numbering starts when a document begins at a nested list level.

- [#175](https://github.com/stella/folio/pull/175) [`d4a40b2`](https://github.com/stella/folio/commit/d4a40b2e28ed37390dc0a8ffbc759ba1d068eb4f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Property-change tracked revisions now resolve fully. Rejecting a `w:pPrChange` restores the stored old paragraph properties wholesale within CT_PPrBase scope (properties the change added are cleared, out-of-scope attrs like the inline `sectPr` still preserved). `w:sectPrChange` and table property changes (`w:tblPrChange`/`w:trPrChange`/`w:tcPrChange`) — previously display-only — now accept and reject: accept keeps live values and clears the record, reject restores the stored previous properties (section rejects keep live header/footer references, which CT_SectPrBase cannot carry). Table property-change records also survive the ProseMirror round-trip instead of being dropped on save, and `acceptAll()`/`rejectAll()` counts include section and table property changes.

- [#158](https://github.com/stella/folio/pull/158) [`bf8e841`](https://github.com/stella/folio/commit/bf8e8411e6e812f7dba5ccb998960e869607507e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fix paragraph-style font-size inheritance for directly formatted DOCX runs.

- [#167](https://github.com/stella/folio/pull/167) [`63eee6c`](https://github.com/stella/folio/commit/63eee6ca8b0188283255c8a6cb69551016b2d5dc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve anchored text boxes hosted by otherwise empty paragraphs.

- [#161](https://github.com/stella/folio/pull/161) [`823b172`](https://github.com/stella/folio/commit/823b172cfa3fcb3b7339fcf60b6705db7b3662cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Ignore collapsible break spaces when deciding whether visible text fits on a line.

- [#146](https://github.com/stella/folio/pull/146) [`58e66c3`](https://github.com/stella/folio/commit/58e66c3ec5e56b7809c8694249c6c69509e9830a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve list continuation, tab-stop precision, and justified hanging-list layout.

- [#120](https://github.com/stella/folio/pull/120) [`c864d1c`](https://github.com/stella/folio/commit/c864d1c0071cb5c7fec1fd59b68aba22def02083) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep hard page breaks distinct from cached Word pagination hints.

- [#157](https://github.com/stella/folio/pull/157) [`7fcb7d7`](https://github.com/stella/folio/commit/7fcb7d7242edbf61dac44aefcad3aa6990a13b71) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Coalesce empty rendered page-break markers after natural paragraph overflow.

- [#124](https://github.com/stella/folio/pull/124) [`411f5e4`](https://github.com/stella/folio/commit/411f5e433c69cb04bc3c2ea4b8d141faa791d80b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce typing latency by coalescing incremental layout on the next animation frame.

- [#126](https://github.com/stella/folio/pull/126) [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Narrow clipboard copy/read failures with an `instanceof Error` check instead of an unchecked `as Error` cast, so `onError` callbacks always receive a real `Error` even when a non-Error value is thrown.

- [#98](https://github.com/stella/folio/pull/98) [`7bdd026`](https://github.com/stella/folio/commit/7bdd026483d29adcc1b57d5916d4e850c2e270d5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Suppress empty hidden list paragraphs and their markers during layout.

- [#95](https://github.com/stella/folio/pull/95) [`50b61d1`](https://github.com/stella/folio/commit/50b61d195834f00c1f07c27b3867a78803e02a91) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve Word parity for tabbed legal paragraphs and splittable keep-next chains.

- [#97](https://github.com/stella/folio/pull/97) [`f0d1cba`](https://github.com/stella/folio/commit/f0d1cba5ca7a1830aae7703e226e731fae76524b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent paragraph-mark character spacing from compressing directly formatted text runs.

- [#127](https://github.com/stella/folio/pull/127) [`f7de365`](https://github.com/stella/folio/commit/f7de36526d276f4c2d5e684da3a468990c075ac4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep default tab stops anchored to the document text-area grid after paragraph indents.

- [#164](https://github.com/stella/folio/pull/164) [`f6d6a7c`](https://github.com/stella/folio/commit/f6d6a7caab18a10b354e60dcae845d77f3a7d8a0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor mirrored left and right margins on facing document pages.

- [#114](https://github.com/stella/folio/pull/114) [`3958226`](https://github.com/stella/folio/commit/3958226cf7047b9b16d506f4e18548ecf23950e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Avoid blank pages when Word's cached page-break hint matches natural paragraph overflow.

- [#142](https://github.com/stella/folio/pull/142) [`1a278bf`](https://github.com/stella/folio/commit/1a278bfe33aeab6f9c4b92923223712d82617547) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve final-section page geometry and clear overflowing first-page headers.

- [#101](https://github.com/stella/folio/pull/101) [`1742abe`](https://github.com/stella/folio/commit/1742abeb2f00f3215767bc3898110463844e279a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor Word rendered-page-break hints during pagination without duplicating structural breaks.

- [#135](https://github.com/stella/folio/pull/135) [`8d9808c`](https://github.com/stella/folio/commit/8d9808c9fd621011b4f73a3bf9ac6e7ee3391a80) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match Word contextual paragraph spacing and cached rendered page-break placement.

- [#165](https://github.com/stella/folio/pull/165) [`6997eee`](https://github.com/stella/folio/commit/6997eee34255bfa28ef3048e7c8bcb443c0ac091) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Apply left-aligned table indents to the first cell text edge.

- [#172](https://github.com/stella/folio/pull/172) [`f1cf28b`](https://github.com/stella/folio/commit/f1cf28bccb5fd5742e45ccdcd651b500c1234950) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fix list marker position for paragraphs with a negative left indent and a hanging indent (`w:ind w:left="-180" w:hanging="360"`). The marker was painted at the left indent instead of `left - hanging`, shifting it one hanging-indent (e.g. 18pt) too far right and mis-indenting continuation lines. The negative left indent is now realized by the line's own `margin-left`, and the marker's remaining negative offset rides on its `margin-left`; positive and zero left-indent lists are unchanged.

- [#103](https://github.com/stella/folio/pull/103) [`3130e5f`](https://github.com/stella/folio/commit/3130e5f616254e6bb127e4c410871c05f6bb386f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add optional DOCX parsing and layout-start performance instrumentation.

- [#173](https://github.com/stella/folio/pull/173) [`2f1a8cb`](https://github.com/stella/folio/commit/2f1a8cbf922b6ddf4fd93f105bbdc95d3ac0faa4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wrap `w:noWrap` table cells whose content overflows a pinned column. When the
  table layout is fixed or an explicit `w:tblW` (`dxa`/`pct`) width pins the
  columns, Word cannot honor `w:noWrap` by widening the column, so it wraps the
  content. Measurement now measures such cells at their real column width (instead
  of an unbounded width), and the painter drops `white-space: nowrap` for them so
  the painted height matches. Auto-width tables still keep `w:noWrap` cells on a
  single line. This corrects under-measured rows that previously let extra rows
  fit per page and dropped a page.

- [#155](https://github.com/stella/folio/pull/155) [`8a3888c`](https://github.com/stella/folio/commit/8a3888c71c6f68970342a6c6cde9af7baf4c1565) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Move oversized keep-with-next chains to a fresh page before paginating them naturally.

- [#133](https://github.com/stella/folio/pull/133) [`d736b57`](https://github.com/stella/folio/commit/d736b57f07953e8faf8793d3324912fa2b8e52f0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce DOCX parsing allocations while converting ordered XML trees.

- [#116](https://github.com/stella/folio/pull/116) [`476fa19`](https://github.com/stella/folio/commit/476fa1918f041d3c9ffcf097b0609d1b0d37b069) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve Aptos document layout and preserve document-default spacing on empty paragraphs.

- [#169](https://github.com/stella/folio/pull/169) [`1fc3f2f`](https://github.com/stella/folio/commit/1fc3f2f748f4330838a56af20f7a3dbf5e5d8959) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Suppress redundant trailing empty paragraph height in populated table cells.

- [#113](https://github.com/stella/folio/pull/113) [`1f74701`](https://github.com/stella/folio/commit/1f74701d3fc2e788a078f8620f006c294c6b8b4e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep named paragraph style fonts on unformatted runs when the paragraph mark has separate formatting.

- [#105](https://github.com/stella/folio/pull/105) [`9697a47`](https://github.com/stella/folio/commit/9697a47eb1c5b93acb374e3ed53e1bdbda47b65e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve zero-sized Word image dimensions during layout conversion.

- [#148](https://github.com/stella/folio/pull/148) [`a5d496f`](https://github.com/stella/folio/commit/a5d496f460c56eec6c7d5c3dd45f528f13f3c144) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Ignore source-less shape placeholders when constructing paragraph image runs.

- [#102](https://github.com/stella/folio/pull/102) [`c01b5b0`](https://github.com/stella/folio/commit/c01b5b0f70ad38c9ff878f8396a418de563ff06a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render WordprocessingGroup drawings as safe SVG previews while preserving their OOXML.

- [#115](https://github.com/stella/folio/pull/115) [`41bc093`](https://github.com/stella/folio/commit/41bc093036e312971532f481d4320e3c63c57f51) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Measure empty paragraphs with their direct paragraph-mark font metrics.

- [#153](https://github.com/stella/folio/pull/153) [`174d0ce`](https://github.com/stella/folio/commit/174d0ce25de076e9a51963547bb37b059b5b21d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve spacing from explicitly selected paragraph styles on empty paragraphs.

- [#171](https://github.com/stella/folio/pull/171) [`f03fa89`](https://github.com/stella/folio/commit/f03fa8966bcee91ed03e33a1bc9f17513960b471) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add horizontal cell borders on top of a table row's `atLeast`/`auto` minimum height instead of absorbing them. When an explicit row height (ECMA-376 §17.4.81 `w:trHeight` without an `hRule`, or `hRule="atLeast"`) exceeds the cell content, the border thickness now extends the row as Word renders it, fixing cumulative vertical drift in tables of short, bordered rows.

- [#162](https://github.com/stella/folio/pull/162) [`0caa2e6`](https://github.com/stella/folio/commit/0caa2e680e0fe125811215ad7ac4a38194ffa673) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Position text-anchored floating tables from their text cursor and keep following text outside full-width table bands.

- [#168](https://github.com/stella/folio/pull/168) [`ffc12e6`](https://github.com/stella/folio/commit/ffc12e6142b03abd4ce27983019b9256719c7b2b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use Montserrat's real line metrics for single-spaced document layout.

- [#119](https://github.com/stella/folio/pull/119) [`4377bf0`](https://github.com/stella/folio/commit/4377bf0db4e7390c409a232852b4d20a533edc08) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Measure collapsed table borders consistently with the table painter.

- [#108](https://github.com/stella/folio/pull/108) [`c7690f9`](https://github.com/stella/folio/commit/c7690f9e77b1a048d8ea6c77676436e3cd191bad) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph run defaults when runs override individual text properties.

- [#137](https://github.com/stella/folio/pull/137) [`1bea7a9`](https://github.com/stella/folio/commit/1bea7a9a4a4ebf8ff263735460a27c2d3ec1241b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match Word's justification of numbered-list continuation lines.

- [#174](https://github.com/stella/folio/pull/174) [`9af69ca`](https://github.com/stella/folio/commit/9af69ca2d63fab2f6795b388184f031b61a87545) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Footnote and endnote body reference marks now display the sequential reference-order number (1, 2, 3…) instead of the raw `w:id`, matching Word for documents with non-contiguous or out-of-order note ids. The body marker and the footnote-area number derive from one shared display-number map, and reserved notes (separators, continuation notices at positive ids) no longer shift numbering.

## 0.4.0

### Minor Changes

- [#90](https://github.com/stella/folio/pull/90) [`b637aa4`](https://github.com/stella/folio/commit/b637aa44b5f6705affa859c037af96b38df4360d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Support opening Agile-encrypted password-protected DOCX files.

### Patch Changes

- [#88](https://github.com/stella/folio/pull/88) [`989999c`](https://github.com/stella/folio/commit/989999cc3629c434a36b16bffee8cba8eb2171b4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fix hidden table row layout, table-cell paragraph spacing, cached field rendering, TOC font inheritance, paragraph-mark caps handling, and parity diagnostics for imported document parity.

- [#93](https://github.com/stella/folio/pull/93) [`63731b6`](https://github.com/stella/folio/commit/63731b61e9be9601adf52a6faceaa7cd1ee9fbc4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve Word parity for legal-template tabs, justified shrink, paragraph spacing collapse, and widow-controlled splits.

- [#91](https://github.com/stella/folio/pull/91) [`521879e`](https://github.com/stella/folio/commit/521879e5b3db5c91e8a68dec0de14e31cc964557) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose the headless watermark API as a public package subpath.

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
