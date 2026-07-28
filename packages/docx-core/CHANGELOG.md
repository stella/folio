# @stll/docx-core

## 0.10.0

### Minor Changes

- [#527](https://github.com/stella/folio/pull/527) [`b4c2b15`](https://github.com/stella/folio/commit/b4c2b1536f1e3483ed0fa55e72c564d543964a41) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project direct paragraph style identifiers and resolved outline levels through the versioned DOCX kernel boundary.

## 0.9.0

### Minor Changes

- [#525](https://github.com/stella/folio/pull/525) [`d0ad1db`](https://github.com/stella/folio/commit/d0ad1db29b9fc6758c77512eeba6b093d539c3b3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project bounded review text and UTF-8/UTF-16 document spans for attributed revisions and comments when OOXML supplies an unambiguous anchor.

## 0.8.0

### Minor Changes

- [#524](https://github.com/stella/folio/pull/524) [`ff873db`](https://github.com/stella/folio/commit/ff873db582719ecf692391bb90054daf25cb0adc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project attributed revisions and comments alongside the document snapshot through one bounded DOCX package call.

### Patch Changes

- [#522](https://github.com/stella/folio/pull/522) [`003ea02`](https://github.com/stella/folio/commit/003ea02f4ac98c29529da7fc4a92bef9d06d9c63) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render OOXML symbol characters with their declared fonts and preserve them through editing.

## 0.7.0

### Minor Changes

- [#520](https://github.com/stella/folio/pull/520) [`e40ccd4`](https://github.com/stella/folio/commit/e40ccd435ea17102fc75910b175e2f78561ac359) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound WordprocessingML scan work and exclude neutral gray highlight colors from semantic spans.

## 0.6.0

### Minor Changes

- [#516](https://github.com/stella/folio/pull/516) [`5ea99cd`](https://github.com/stella/folio/commit/5ea99cd32d44db94afac8bb44c31c8f32bc2aa19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a browser-native DOCX projection API backed by a bounded Rust WebAssembly kernel.

### Patch Changes

- [#519](https://github.com/stella/folio/pull/519) [`b003927`](https://github.com/stella/folio/commit/b003927467052dcc6c6c2c3ddd66cebf057e7f84) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve typed errors for invalid WordprocessingML text placement and document type declarations.

## 0.5.2

### Patch Changes

- [#497](https://github.com/stella/folio/pull/497) [`3b28632`](https://github.com/stella/folio/commit/3b28632025d3798b4bd4b9c8268fbb444237ce6c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve Word positional tabs across parsing, editing, and serialization.

- [#502](https://github.com/stella/folio/pull/502) [`8b2535e`](https://github.com/stella/folio/commit/8b2535ec268eba128895251a3d573af29c241b16) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve localized comment annotation-reference formatting across DOCX serialization.

- [#493](https://github.com/stella/folio/pull/493) [`b06a26d`](https://github.com/stella/folio/commit/b06a26d6efeb818d12ed78799a626f5d058494e8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve DrawingML gradient fill details across text boxes and other shared fill consumers.

- [#494](https://github.com/stella/folio/pull/494) [`0a613b8`](https://github.com/stella/folio/commit/0a613b879914f40ea2fc143caaed54fa2bd9412e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored DrawingML outline details across shapes and text boxes.

## 0.5.1

### Patch Changes

- [#470](https://github.com/stella/folio/pull/470) [`db702de`](https://github.com/stella/folio/commit/db702dea23ca7f7374031c13aac64ad2f520d981) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve Word font script hints across document edits and saves.

- [#472](https://github.com/stella/folio/pull/472) [`76c46d5`](https://github.com/stella/folio/commit/76c46d53f980c1651c83bd600bfbd565822cfa53) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve distinct image names, descriptions, and titles when saving DOCX files.

- [#461](https://github.com/stella/folio/pull/461) [`9ce81df`](https://github.com/stella/folio/commit/9ce81dffee6ed812b91ee7a5cdb839c5d2d0f690) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve clickable image hyperlinks across DOCX save and reopen.

## 0.5.0

### Minor Changes

- [#412](https://github.com/stella/folio/pull/412) [`21274be`](https://github.com/stella/folio/commit/21274be83afaadc9d28053c87b5ea84ea619c491) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a first-class suggestion layer to tracked changes. AI-proposed edits can be
  applied with the new `"suggested"` apply mode: they render with the
  tracked-change grammar but a dotted stroke and a dedicated hue, and are always
  stripped from serialized DOCX output until accepted. Accepting a suggestion
  converts it into a normal tracked change authored by the accepting user (or, for
  a whole inserted table, applies it directly since OOXML has no tracked
  representation for it); rejecting inverse-applies it.

  Suggested mode covers inline text/format operations (`replaceInBlock`,
  `replaceRange`, `formatRange`) and block/table structural operations
  (`insertAfterBlock`, `insertBeforeBlock`, `replaceBlock`, `deleteBlock`,
  `insertSignatureTable`, `insertTableRow`, `deleteTableRow`, `insertTableColumn`,
  `deleteTableColumn`). Whole-node inserts are stripped entirely; suggested
  deletes serialize as though they never happened; the strip is the single
  `fromProseDoc`/`extractBlocks` boundary every serialization path funnels through.
  Cell merge/split and comment operations remain `unsupportedMode`.

  New core commands (`getSuggestions`, `acceptSuggestion`, `acceptAllSuggestions`,
  `rejectSuggestion`, `rejectAllSuggestions`, `findSuggestionRange`) and
  editor-ref methods (`getSuggestions`, `acceptSuggestion` returning
  `{ accepted, appliedAs }`, `rejectSuggestion`, `scrollToSuggestion`) expose the
  layer to hosts; `getSuggestions` reports each suggestion's kinds and `appliedAs`
  (`"tracked"` vs `"direct"`). The React and Vue adapters expose the same ref
  surface (the Nuxt module re-exports it).

  Tracked changes also gain an optional `initials` field, carried through the
  model and the ProseMirror marks/node attrs for UI attribution (hover, accept
  authoring). It is intentionally NOT serialized onto `w:ins`/`w:del`/`w:*PrChange`
  or table row/cell markers — `w:initials` is not part of ECMA-376
  `CT_TrackChange`, so output stays schema-strict — but the parser remains tolerant
  of it if an external document supplies one.

### Patch Changes

- [#427](https://github.com/stella/folio/pull/427) [`64f0737`](https://github.com/stella/folio/commit/64f07378ba3f460b999a8a7bba822ed0a01e37e0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep tracked-change revision ids within the range OOXML consumers accept. Suggestion-mode edits seeded their `w:ins`/`w:del` id counter from the clock, producing 13-digit `w:id` values that made exported documents fail to open. Ids now continue from the document's own highest revision id. Port of eigenpal/docx-editor#1093.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound version-comparison, agent, and validation paths against crafted-input
  resource exhaustion: a single aggregate LCS cell budget is now shared across all
  document stories; move detection dequeues in O(1) instead of `Array.shift`;
  `diffWordSegments` caps its DP matrix and falls back to a whole-string diff;
  `ensureParaIds` and `docx-core`'s `validateDocxPackage` enforce entry-count and
  uncompressed-size limits before reading; note-paragraph patching builds a linear
  offset index instead of rescanning per id; agent whole-word search uses a bounded
  boundary window; `suggest_changes` enforces an aggregate operation-text budget;
  and tracked vertical cell split refuses a stored continuation whose `gridSpan`
  exceeds one column.

## 0.4.0

### Minor Changes

- [#358](https://github.com/stella/folio/pull/358) [`a96f6e5`](https://github.com/stella/folio/commit/a96f6e51908e7f04955240763f1e198bdd38f374) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and edit tables inside text boxes.

- [#384](https://github.com/stella/folio/pull/384) [`f349951`](https://github.com/stella/folio/commit/f34995146f0ee2a7838a6cc9c501e0227b9b1250) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add vertical cell merge revision import, review, resolution, and export support.

### Patch Changes

- [#414](https://github.com/stella/folio/pull/414) [`d4d51c6`](https://github.com/stella/folio/commit/d4d51c627f605e9d1e335402cc3556da404e4847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve cached OOXML page boundaries inside paragraphs during editing and layout.

- [#374](https://github.com/stella/folio/pull/374) [`c478c54`](https://github.com/stella/folio/commit/c478c540eade004a1bffbc518b29191bb18ed7d9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve text boxes nested inside inline DOCX content controls, including tracked moves, through editing and save.

- [#389](https://github.com/stella/folio/pull/389) [`f3d2847`](https://github.com/stella/folio/commit/f3d284783043e162fd0e2d006dbbcdfae5b0969e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor legacy OOXML compatibility modes when fitting justified lines.

- [#398](https://github.com/stella/folio/pull/398) [`0b73404`](https://github.com/stella/folio/commit/0b73404618886852439a57f8d9c257a0a448709c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor section line grids, paragraph grid opt-outs, hidden table-cell markers, and explicit zero cell margins during layout.

- [#367](https://github.com/stella/folio/pull/367) [`0dd5214`](https://github.com/stella/folio/commit/0dd5214f26bdc6a82a9273290b004cbf5fee43bc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicitly empty DOCX comment authors during parsing and serialization.

## 0.3.0

### Minor Changes

- [#303](https://github.com/stella/folio/pull/303) [`689dbf5`](https://github.com/stella/folio/commit/689dbf553a028864fef280b5773eeff0fbe40d26) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and render both OOXML table-cell diagonal border directions.

- [#309](https://github.com/stella/folio/pull/309) [`1ec610f`](https://github.com/stella/folio/commit/1ec610f362aab68fc55807edef88974304c22bf4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add locale-aware DOCX automatic hyphenation and tighter Word hanging-punctuation layout.

- [#287](https://github.com/stella/folio/pull/287) [`fbc7fce`](https://github.com/stella/folio/commit/fbc7fce4c977eabace64a2756c51e42e788c5370) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add replaceable Unicode line breaking with DOCX language, kinsoku, custom line-edge, and compatibility-rule support.

### Patch Changes

- [#308](https://github.com/stella/folio/pull/308) [`f56c68c`](https://github.com/stella/folio/commit/f56c68c2b9a1f7c03186617da6da869e80c4e187) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit bold formatting on automatic list markers through parsing, layout, and painting.

- [#311](https://github.com/stella/folio/pull/311) [`482e5e7`](https://github.com/stella/folio/commit/482e5e787f226a552bbe272d0816561ba9389877) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor break-only paragraph placement and exact image-only line footprints during pagination.

- [#328](https://github.com/stella/folio/pull/328) [`166db7f`](https://github.com/stella/folio/commit/166db7fe854ddaac94c7739c7c64caa601313027) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve tracked insertions and deletions inside inline DOCX content controls.

## 0.2.0

### Minor Changes

- [#272](https://github.com/stella/folio/pull/272) [`af10f08`](https://github.com/stella/folio/commit/af10f0840565680b087fd2955ba2ab7c512e628f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add DOCX conformance detection to parsed and created packages.

### Patch Changes

- [#237](https://github.com/stella/folio/pull/237) [`f82a489`](https://github.com/stella/folio/commit/f82a489b0af3f18cdcd226b2e7b10074c5ce80b1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve shape-to-text fitting for imported OOXML text boxes.

- [#244](https://github.com/stella/folio/pull/244) [`ef9b7a6`](https://github.com/stella/folio/commit/ef9b7a6797ab288bf446265e71c87abf759e08fc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve OOXML paragraph-frame spacing, keep drop caps in normal flow, and retain side wrapping for single frames.

- [#271](https://github.com/stella/folio/pull/271) [`06adc8d`](https://github.com/stella/folio/commit/06adc8dff272e2272a87c28886f456b8c74e1bd0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve omitted table grid columns and use them when measuring row cells.

- [#241](https://github.com/stella/folio/pull/241) [`becca9c`](https://github.com/stella/folio/commit/becca9c26cb4f032c18731e2c5b412461c5dd85c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve numbering-level marker alignment through parsing and layout.

## 0.1.1

### Patch Changes

- [#201](https://github.com/stella/folio/pull/201) [`46c6730`](https://github.com/stella/folio/commit/46c6730ebf29daccdfac64c72fcf07702709e70f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Publish the canonical DOCX model, legal-source compiler, serializer, and validator from folio.
