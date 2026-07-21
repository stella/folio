# @stll/folio-core

## 0.14.0

### Minor Changes

- [#433](https://github.com/stella/folio/pull/433) [`45af3e8`](https://github.com/stella/folio/commit/45af3e8073136efbf65d13881e8d7420d2402600) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `docxToMarkdown(bytes)` to the `/server` entry: a one-call, server-safe DOCX
  bytes → markdown converter that composes `parseDocx` (font preloading disabled)
  and `toMarkdown`, so non-browser callers get full DOCX fidelity without deep
  imports or a hand-rolled OOXML walker. Also fixes the DOCX table-cell parser to
  descend into block-level content controls (`w:tc > w:sdt > w:sdtContent`), so
  controlled field text inside table cells is no longer dropped.

### Patch Changes

- [#431](https://github.com/stella/folio/pull/431) [`34b0737`](https://github.com/stella/folio/commit/34b0737e6e110d9f2aa464a4f40ad13aead5ceeb) Thanks [@cursor](https://github.com/apps/cursor)! - Honor tblpXSpec over tblpX for floating tables, and keep bar tab stops from suppressing the default tab grid.

## 0.13.0

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

- [#422](https://github.com/stella/folio/pull/422) [`75842cf`](https://github.com/stella/folio/commit/75842cf60c290af3f756e7dbea7f95671fbdea4f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match Word's final-line paragraph-mark spacing for visible lists.

- [#428](https://github.com/stella/folio/pull/428) [`ce930f4`](https://github.com/stella/folio/commit/ce930f4ee45d2b793ef0d625fb0598ce008cb600) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Drop unresolvable comment range and reference markers during export, porting eigenpal [#1090](https://github.com/stella/folio/issues/1090).

- [#429](https://github.com/stella/folio/pull/429) [`4b6e885`](https://github.com/stella/folio/commit/4b6e88531408fc9ecb82ae7c0a71e797864fa996) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Port eigenpal/docx-editor#1096 image border rendering through layout painting.

- [#422](https://github.com/stella/folio/pull/422) [`75842cf`](https://github.com/stella/folio/commit/75842cf60c290af3f756e7dbea7f95671fbdea4f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Ignore cached table page boundaries when tracked row content can change pagination.

- [#427](https://github.com/stella/folio/pull/427) [`64f0737`](https://github.com/stella/folio/commit/64f07378ba3f460b999a8a7bba822ed0a01e37e0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep tracked-change revision ids within the range OOXML consumers accept. Suggestion-mode edits seeded their `w:ins`/`w:del` id counter from the clock, producing 13-digit `w:id` values that made exported documents fail to open. Ids now continue from the document's own highest revision id. Port of eigenpal/docx-editor#1093.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Harden agent operation integrity: `read_document`/`read_section`/`find_text`
  now return a per-block text hash and `suggest_changes`/`add_comment` accept a
  caller-supplied precondition, so a stale edit prepared against content the model
  read earlier is detected instead of being stamped from a fresh apply-time
  snapshot. Operation-mode and block-range lookups use own-property checks so
  prototype keys (`__proto__`, `constructor`) can no longer crash the API.

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

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent untrusted DOCX assets from affecting the host page. Embedded fonts are
  registered under per-document scoped family names (resolved through the font
  resolver) so a document embedding a face named after a host UI family can no
  longer shadow it page-wide, and watermark dialogs validate external image
  targets against an http/https allowlist (with a defensive guard before emitting
  an external relationship) so `file:`/UNC/other-scheme targets cannot be written
  into exported documents.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop hidden and active DOCX content from leaking through save and read paths.
  Selective save now bails to a full repack when the package contains
  non-preservable entries (e.g. `word/vbaProject.bin`, embedded binaries) instead
  of round-tripping them; hidden table-row text and `w:vanish` runs are excluded
  from the AI snapshot; footnotes referenced only from hidden rows are no longer
  painted; metadata-privacy scrubbing matches `docProps/core.xml` case-insensitively;
  server text extraction resolves referenced headers/footers via relationships
  instead of reading orphan parts; bound content-control clicks no longer throw;
  and text-box anchor markers are stripped from pasted HTML and salted with a
  per-load nonce.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent CSS and OOXML injection from attacker-controlled document/collaboration
  values. Colors are validated to a strict hex/`auto` format at a single
  `colorResolver` choke point (closing themed table-fill and diagonal-border
  `url()` injection and the pasted `data-bgcolor` path); comment `paraId`/`textId`
  are validated to 8-hex at parse and XML-escaped on serialize; run/paragraph/
  table/style color and theme attributes are XML-escaped and hex-validated; inline
  and block SDT raw properties are replayed only when they are a single
  well-formed `w:sdtPr`/`w:sdtEndPr` element (otherwise synthesized); remote
  collaborator colors are validated before use and painted via `backgroundColor`
  (not the `background` shorthand); and controlled comments are sanitized before
  becoming editor state.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound DOCX parsing and layout against crafted-input resource exhaustion: clamp
  table `gridSpan`/column counts, section column count, and page dimensions;
  floor the default tab stop; cap kinsoku/line-break rule lists (stored as sets)
  and run language tags; add an iteration cap to cross-run hyphenation; replace
  the conformance root-tag regex with a linear, non-backtracking scanner; cap
  per-element xmlns declarations; enforce an incremental size budget while
  building grouped-drawing SVG previews; make style-numbered list resume O(1) per
  paragraph; and guard encrypted-DOCX parsing with DIFAT cycle detection plus a
  `spinCount` ceiling.

- [#426](https://github.com/stella/folio/pull/426) [`3229068`](https://github.com/stella/folio/commit/32290689f2256e2e601f1be6701aceb5d135169f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Harden untrusted clipboard HTML, image sources, and VML style parsing against XSS, remote fetch, ReDoS, and prototype pollution.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Sanitize hyperlink and image link URLs so pasted, programmatic, or DOCX-sourced
  `javascript:`/`data:`/`file:` targets can no longer reach the live DOM or be
  opened. Hyperlink marks are sanitized on parse (`parseDOM`), on render
  (`toDOM`), and when set/inserted/edited; image `a:hlinkClick` targets are
  sanitized at parse time; the Vue popup `window.open` path now mirrors React's
  sanitizer; and aux-click on link anchors no longer bypasses the guard. Internal
  bookmark anchors (`#name`) are preserved.
- Updated dependencies [[`64f0737`](https://github.com/stella/folio/commit/64f07378ba3f460b999a8a7bba822ed0a01e37e0), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`21274be`](https://github.com/stella/folio/commit/21274be83afaadc9d28053c87b5ea84ea619c491)]:
  - @stll/docx-core@0.5.0

## 0.12.0

### Minor Changes

- [#397](https://github.com/stella/folio/pull/397) [`0056754`](https://github.com/stella/folio/commit/0056754d22202cf27e9cd734305b10e24cc8d899) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add shared editable footnote and endnote stories with matching React and Vue surfaces.

- [#392](https://github.com/stella/folio/pull/392) [`f04bde1`](https://github.com/stella/folio/commit/f04bde1f25af45e7155c469255c900b242cd1f32) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render configurable page margin guides in both framework adapters.

- [#358](https://github.com/stella/folio/pull/358) [`a96f6e5`](https://github.com/stella/folio/commit/a96f6e51908e7f04955240763f1e198bdd38f374) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and edit tables inside text boxes.

- [#362](https://github.com/stella/folio/pull/362) [`49a610a`](https://github.com/stella/folio/commit/49a610a89be5e4e479557289f973111ff6ec530c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render tables inside text boxes through the layout pipeline.

- [#384](https://github.com/stella/folio/pull/384) [`f349951`](https://github.com/stella/folio/commit/f34995146f0ee2a7838a6cc9c501e0227b9b1250) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add vertical cell merge revision import, review, resolution, and export support.

- [#407](https://github.com/stella/folio/pull/407) [`7edee13`](https://github.com/stella/folio/commit/7edee13caa2b6f5fca6defdf9a980f4ea4a81f73) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a versioned DOCX conformance report to the server API.

- [#417](https://github.com/stella/folio/pull/417) [`8359b99`](https://github.com/stella/folio/commit/8359b9949eed5d56820217e18432666c5dfd2d5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add guarded XML patch proposal evaluation for server workflows.

- [#388](https://github.com/stella/folio/pull/388) [`4f9c04e`](https://github.com/stella/folio/commit/4f9c04ec03a94bb25cfab3810d752470ddf18e01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add tracked vertical table cell split operations with reversible native revisions and agent-tool support.

- [#360](https://github.com/stella/folio/pull/360) [`908861b`](https://github.com/stella/folio/commit/908861ba38d1277d39bd50457dcdea14f0b45b91) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fire copy, cut, and paste callbacks consistently from the shared hidden editor in both adapters.

- [#371](https://github.com/stella/folio/pull/371) [`0c7f14d`](https://github.com/stella/folio/commit/0c7f14d48b9cddfa2ba77e79880d40962d91a9e7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add direct table column deletion operations.

- [#366](https://github.com/stella/folio/pull/366) [`9a7fe2f`](https://github.com/stella/folio/commit/9a7fe2f7d704ab023ac852250bd96c1c124af714) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a stable document operation for direct table-row deletion.

- [#368](https://github.com/stella/folio/pull/368) [`18c8fdf`](https://github.com/stella/folio/commit/18c8fdf1cc5deb7df5d37f4467e07c115c1a0275) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a stable document operation for direct table-column insertion.

- [#364](https://github.com/stella/folio/pull/364) [`e3d7759`](https://github.com/stella/folio/commit/e3d7759dd39e06e79823e647c6b313d178f38f5d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a stable document operation for direct table-row insertion.

- [#373](https://github.com/stella/folio/pull/373) [`0ada098`](https://github.com/stella/folio/commit/0ada0989e2781966b3296027e6c4ec8e5e45383b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add direct rectangular table cell merging.

- [#405](https://github.com/stella/folio/pull/405) [`d20fb9f`](https://github.com/stella/folio/commit/d20fb9fc7ae0777077e7daf5931539e63a6672cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `highlightPassage` / `clearPassageHighlight` to `DocxEditorRef`: resolve `{ blockId, text }` to a range inside the block, scroll to it, and paint a persistent translucent passage highlight, falling back to scroll-to-block with a paragraph flash when the text no longer matches. Core exports the framework-neutral `resolvePassageRange` resolver.

- [#357](https://github.com/stella/folio/pull/357) [`997f0c1`](https://github.com/stella/folio/commit/997f0c12d4418239ebd86e9b20d2058f71cbf0e1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share content-control picker state and dispatch across adapters, and add interactive Vue dropdown and date controls.

- [#375](https://github.com/stella/folio/pull/375) [`460e1c3`](https://github.com/stella/folio/commit/460e1c346c025b04a326cb243d20b2172f42265a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add direct table cell splitting.

- [#421](https://github.com/stella/folio/pull/421) [`740b282`](https://github.com/stella/folio/commit/740b2822a9e52276b324a9c14982a0e9bebea32a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add guarded XML patch application with complete output validation.

- [#408](https://github.com/stella/folio/pull/408) [`5baa329`](https://github.com/stella/folio/commit/5baa3292273d72db9279f5e335aed9f620c20229) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add bounded read-only DOCX package inspection.

- [#387](https://github.com/stella/folio/pull/387) [`a56ce7d`](https://github.com/stella/folio/commit/a56ce7d0018948e1e78a3139a00577e3c7858d09) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Support tracked inline formatting operations across document review surfaces.

- [#381](https://github.com/stella/folio/pull/381) [`f6c6312`](https://github.com/stella/folio/commit/f6c63128e3b0aceab7a473ed55b98a8a1be79ed0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and resolve tracked table cell insertions and deletions across review APIs.

- [#383](https://github.com/stella/folio/pull/383) [`a40ada1`](https://github.com/stella/folio/commit/a40ada166587eff96136db8d7404480451b43e0f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add tracked table column deletion across document operations and agent tools.

- [#382](https://github.com/stella/folio/pull/382) [`4bd9918`](https://github.com/stella/folio/commit/4bd99187c88c38e903ea95d617aea77db664c3b1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add tracked table column insertion across document operations and agent tools.

- [#380](https://github.com/stella/folio/pull/380) [`3cc5128`](https://github.com/stella/folio/commit/3cc512892017afa96b44a0cb609f6f3d81ceeb7b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add tracked table row deletion with review resolution and agent-tool support.

- [#377](https://github.com/stella/folio/pull/377) [`3eb34b4`](https://github.com/stella/folio/commit/3eb34b41acaffd84ac85afaf144e3dc78afebb4c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add tracked table row insertion with review resolution and agent-tool support.

- [#391](https://github.com/stella/folio/pull/391) [`5fc68b8`](https://github.com/stella/folio/commit/5fc68b87e4362a726428a944d232b5417ca4dda3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add tracked vertical table cell merge operations with row-count targeting and reversible native revisions.

### Patch Changes

- [#385](https://github.com/stella/folio/pull/385) [`ba753fb`](https://github.com/stella/folio/commit/ba753fbb07ba71370216975962002bb59c569769) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor authored no-break spaces without inventing locale-specific wrapping.

- [#414](https://github.com/stella/folio/pull/414) [`d4d51c6`](https://github.com/stella/folio/commit/d4d51c627f605e9d1e335402cc3556da404e4847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve cached OOXML page boundaries inside paragraphs during editing and layout.

- [#374](https://github.com/stella/folio/pull/374) [`c478c54`](https://github.com/stella/folio/commit/c478c540eade004a1bffbc518b29191bb18ed7d9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve text boxes nested inside inline DOCX content controls, including tracked moves, through editing and save.

- [#409](https://github.com/stella/folio/pull/409) [`3ca97c9`](https://github.com/stella/folio/commit/3ca97c913741711e91c24cdee46017d04802bb01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve theme fonts during layout and preserve paragraph style fonts inside table styles.

- [#389](https://github.com/stella/folio/pull/389) [`f3d2847`](https://github.com/stella/folio/commit/f3d284783043e162fd0e2d006dbbcdfae5b0969e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor legacy OOXML compatibility modes when fitting justified lines.

- [#415](https://github.com/stella/folio/pull/415) [`b39aa26`](https://github.com/stella/folio/commit/b39aa2689c98ee688a9d59219c283832ff0abecd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor cached rendered page boundaries at table-row starts when a row would otherwise split.

- [#402](https://github.com/stella/folio/pull/402) [`930ffc8`](https://github.com/stella/folio/commit/930ffc8dc06604bd5279405f8d20bb720dc9673f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route formatting, focus, and history commands to the active document story consistently.

- [#376](https://github.com/stella/folio/pull/376) [`354b449`](https://github.com/stella/folio/commit/354b449f7f74f182c02b69eacb6c8eaaa193f483) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route painted content-control clicks through the shared widget controller and render usable pickers in React and Vue.

- [#398](https://github.com/stella/folio/pull/398) [`0b73404`](https://github.com/stella/folio/commit/0b73404618886852439a57f8d9c257a0a448709c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor section line grids, paragraph grid opt-outs, hidden table-cell markers, and explicit zero cell margins during layout.

- [#406](https://github.com/stella/folio/pull/406) [`ed40fac`](https://github.com/stella/folio/commit/ed40fac565fcbe264242d5e404a8cd0ad6e3f813) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor document grid types, preserve section marks after trailing page breaks, and reconcile cached page boundaries after tables.

- [#379](https://github.com/stella/folio/pull/379) [`4852343`](https://github.com/stella/folio/commit/4852343a771cd78540d8fb3d79f92fa0b988cc12) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve bounded space contraction for justified hanging list continuations.

- [#418](https://github.com/stella/folio/pull/418) [`83434c3`](https://github.com/stella/folio/commit/83434c3e4ea9bd8e75e8532c892dfe785e5ac3a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use a closer bundled fallback for Aptos document fonts

- [#416](https://github.com/stella/folio/pull/416) [`0f619d4`](https://github.com/stella/folio/commit/0f619d4404d6a09d99850105fbad62dafa06bef7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Measure and paint ordinary DOCX bold text with the standard 700 font weight.

- [#404](https://github.com/stella/folio/pull/404) [`16fc2e1`](https://github.com/stella/folio/commit/16fc2e18506de34f3e5cde3cf70885c1e300d40d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Avoid reparsing the complete document XML when detecting DOCX conformance.

- [#411](https://github.com/stella/folio/pull/411) [`01a4dc0`](https://github.com/stella/folio/commit/01a4dc083f439fa4970a5ea03cdfa7f47e212e9f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve automatic line metrics and cached page-boundary reconciliation.

- [#365](https://github.com/stella/folio/pull/365) [`1092005`](https://github.com/stella/folio/commit/1092005cea1b227febdd833bd12b1ff9662b3c31) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve implicit default paragraph style spacing on empty paragraphs.

- [#369](https://github.com/stella/folio/pull/369) [`9836577`](https://github.com/stella/folio/commit/983657797c25746f8eeae01a26fce5cc6b2ca352) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve tracked-change wrappers around text boxes through editing and save.

- [#386](https://github.com/stella/folio/pull/386) [`43a68c9`](https://github.com/stella/folio/commit/43a68c9c3a989683f825a5756b8fd036c7549f60) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inline text box source order across document edits.

- [#410](https://github.com/stella/folio/pull/410) [`5f888df`](https://github.com/stella/folio/commit/5f888df00a4c11b317360d2419c0b9deabdfcce3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce DOCX parsing and ProseMirror conversion work for large documents.

- [#413](https://github.com/stella/folio/pull/413) [`2e73ef5`](https://github.com/stella/folio/commit/2e73ef56f45ed7161f14447d76a63a937372642f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Limit hanging punctuation to East Asian language runs so Latin punctuation stays within line bounds.

- [#367](https://github.com/stella/folio/pull/367) [`0dd5214`](https://github.com/stella/folio/commit/0dd5214f26bdc6a82a9273290b004cbf5fee43bc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicitly empty DOCX comment authors during parsing and serialization.

- [#385](https://github.com/stella/folio/pull/385) [`ba753fb`](https://github.com/stella/folio/commit/ba753fbb07ba71370216975962002bb59c569769) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refine bounded first-line contraction for justified deep-hanging list markers.

- [#389](https://github.com/stella/folio/pull/389) [`f3d2847`](https://github.com/stella/folio/commit/f3d284783043e162fd0e2d006dbbcdfae5b0969e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use paragraph-mark metrics for whitespace-only paragraphs.

- [#385](https://github.com/stella/folio/pull/385) [`ba753fb`](https://github.com/stella/folio/commit/ba753fbb07ba71370216975962002bb59c569769) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Paint OOXML no-break hyphens with the ordinary hyphen glyph.

- [#394](https://github.com/stella/folio/pull/394) [`3ce8392`](https://github.com/stella/folio/commit/3ce83920b8d706dd0e2084967e9705b155914a0a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph fonts when table styles add run formatting.

- [#396](https://github.com/stella/folio/pull/396) [`c768500`](https://github.com/stella/folio/commit/c768500d27bacc495a154d1e9d0771cc2638d669) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Speed up paragraph line breaking, text measurement caching, and initial editor state construction.

- Updated dependencies [[`a96f6e5`](https://github.com/stella/folio/commit/a96f6e51908e7f04955240763f1e198bdd38f374), [`d4d51c6`](https://github.com/stella/folio/commit/d4d51c627f605e9d1e335402cc3556da404e4847), [`f349951`](https://github.com/stella/folio/commit/f34995146f0ee2a7838a6cc9c501e0227b9b1250), [`c478c54`](https://github.com/stella/folio/commit/c478c540eade004a1bffbc518b29191bb18ed7d9), [`f3d2847`](https://github.com/stella/folio/commit/f3d284783043e162fd0e2d006dbbcdfae5b0969e), [`0b73404`](https://github.com/stella/folio/commit/0b73404618886852439a57f8d9c257a0a448709c), [`0dd5214`](https://github.com/stella/folio/commit/0dd5214f26bdc6a82a9273290b004cbf5fee43bc)]:
  - @stll/docx-core@0.4.0

## 0.11.0

### Minor Changes

- [#353](https://github.com/stella/folio/pull/353) [`10f395e`](https://github.com/stella/folio/commit/10f395e216668f4a6779aa652f0d0b783e838f31) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share find contracts and matching logic across adapters and add a functional Vue find/replace binding.

- [#355](https://github.com/stella/folio/pull/355) [`f2abf34`](https://github.com/stella/folio/commit/f2abf34f20e32d1094f4459ab633720e9712e439) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share persistent header and footer editing across the React and Vue adapters.

- [#354](https://github.com/stella/folio/pull/354) [`fa23331`](https://github.com/stella/folio/commit/fa233311ff79811818ce11bab7f9854097698c93) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share collaboration module loading and remote-selection painting, and complete the Vue collaboration pipeline.

- [#350](https://github.com/stella/folio/pull/350) [`860a894`](https://github.com/stella/folio/commit/860a894e4ffe94c1492a04db8aa7fc9736f097df) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add shared rendered-DOM geometry, functional Vue selection and decoration overlays, and document-independent table cell tracking.

## 0.10.0

### Minor Changes

- [#343](https://github.com/stella/folio/pull/343) [`fd5717b`](https://github.com/stella/folio/commit/fd5717b3465d12dd114467106ae552e10df96699) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add auditable privacy transforms for version comparison output.

- [#340](https://github.com/stella/folio/pull/340) [`5b99d1e`](https://github.com/stella/folio/commit/5b99d1ed01d0c10cecce2850dc611fb336e3f725) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add immutable reviewed-view projections for document stories.

- [#341](https://github.com/stella/folio/pull/341) [`21b7e0e`](https://github.com/stella/folio/commit/21b7e0e4ba6be4303cd351195e44e0ed305de105) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compare document stories with source-specific block handles.

- [#342](https://github.com/stella/folio/pull/342) [`8f8b569`](https://github.com/stella/folio/commit/8f8b569f29a2c5f78001e12c3536d8aa7b4227a5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add selectable text, formatting, and metadata comparison scopes.

- [#338](https://github.com/stella/folio/pull/338) [`b324ce3`](https://github.com/stella/folio/commit/b324ce3154899d2e311529612d7855840b64afcc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add story-scoped document operations for footnotes and endnotes.

- [#347](https://github.com/stella/folio/pull/347) [`d2e0c4a`](https://github.com/stella/folio/commit/d2e0c4a774381f5daae028839de9f639b8de2ad6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add auditable package-metadata privacy rewriting that persists through later saves.

- [#349](https://github.com/stella/folio/pull/349) [`692528b`](https://github.com/stella/folio/commit/692528b42a3da66e425497cb4075f3cd82609524) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add package-metadata privacy options and reports to redline generation.

- [#346](https://github.com/stella/folio/pull/346) [`b2f3885`](https://github.com/stella/folio/commit/b2f3885619c85a29f54a74be13a7d89a6dbee46c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate tracked changes across matched document stories with selectable resolved input views.

### Patch Changes

- [#333](https://github.com/stella/folio/pull/333) [`b77268b`](https://github.com/stella/folio/commit/b77268bb4cb4d57f2c892138ddeba97d8f16b028) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor table-cell positioning scope for anchored images.

- [#336](https://github.com/stella/folio/pull/336) [`4a4b10f`](https://github.com/stella/folio/commit/4a4b10ff5b12cdeec66374f32daed84794bf5b9c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Tighten justified paragraph fitting using common-layout reference endpoints.

- [#344](https://github.com/stella/folio/pull/344) [`d6df33d`](https://github.com/stella/folio/commit/d6df33d7b8fb536cdc4c8125dfbdd39fbd06ef49) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Default omitted paragraph hanging-punctuation settings to enabled during line layout.

- [#339](https://github.com/stella/folio/pull/339) [`2337ae1`](https://github.com/stella/folio/commit/2337ae1b0548982775a93600260adb2f0aaa3ecc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor the document hyphenation zone when choosing automatic line breaks.

## 0.9.0

### Minor Changes

- [#304](https://github.com/stella/folio/pull/304) [`d3c2816`](https://github.com/stella/folio/commit/d3c2816d2bd48d04fea6abe4924d3c93d7d0104c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove the static DOCX capability catalog while retaining document-specific diagnostics.

- [#303](https://github.com/stella/folio/pull/303) [`689dbf5`](https://github.com/stella/folio/commit/689dbf553a028864fef280b5773eeff0fbe40d26) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and render both OOXML table-cell diagonal border directions.

- [#331](https://github.com/stella/folio/pull/331) [`59a581a`](https://github.com/stella/folio/commit/59a581a7f530e9d1cdc9d35702ad54f1cb5bb2a6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add story-scoped document operations for headers and footers.

- [#309](https://github.com/stella/folio/pull/309) [`1ec610f`](https://github.com/stella/folio/commit/1ec610f362aab68fc55807edef88974304c22bf4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add locale-aware DOCX automatic hyphenation and tighter Word hanging-punctuation layout.

- [#306](https://github.com/stella/folio/pull/306) [`9972201`](https://github.com/stella/folio/commit/99722013cdbf2320a0ca90e673a2716d12c91030) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add portable document style sets, the Stella Style drafting preset, sanitized DOCX style extraction, and complete fresh-document style serialization.

- [#287](https://github.com/stella/folio/pull/287) [`fbc7fce`](https://github.com/stella/folio/commit/fbc7fce4c977eabace64a2756c51e42e788c5370) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add replaceable Unicode line breaking with DOCX language, kinsoku, custom line-edge, and compatibility-rule support.

### Patch Changes

- [#301](https://github.com/stella/folio/pull/301) [`2b6f212`](https://github.com/stella/folio/commit/2b6f212ac4aa776c6ef61b51596c4e37e3154dd0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render bottom-to-top table cell text in paged layout.

- [#314](https://github.com/stella/folio/pull/314) [`8bd16d0`](https://github.com/stella/folio/commit/8bd16d09599a2bf99f070c5a0ed686db97709d04) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Position vertical OOXML image crops relative to the bitmap height.

- [#283](https://github.com/stella/folio/pull/283) [`79b74e3`](https://github.com/stella/folio/commit/79b74e3a37743cdc1f05af8be11b3c0c2bc3e03e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Base justified list marker wrapping on measured compressible-space width.

- [#274](https://github.com/stella/folio/pull/274) [`fc01a2d`](https://github.com/stella/folio/commit/fc01a2d1857a9d67378e2441560998c413278397) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render relationship-backed previews for legacy embedded objects with their authored line height.

- [#312](https://github.com/stella/folio/pull/312) [`9f0ddce`](https://github.com/stella/folio/commit/9f0ddce46fb7541af2640dc42932fc789cdf726e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored table-cell paragraphs and keep paragraph-mark fonts off existing body text.

- [#186](https://github.com/stella/folio/pull/186) [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accept DOCX comments with missing or empty author metadata.

- [#315](https://github.com/stella/folio/pull/315) [`e476fda`](https://github.com/stella/folio/commit/e476fda883fbcda84b243bb144113d2bc47ef0ff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Hyphenate words across adjacent inline formatting runs to match Word line endpoints.

- [#299](https://github.com/stella/folio/pull/299) [`c9687eb`](https://github.com/stella/folio/commit/c9687ebd0972a894ea7c96fac7331fd8702315a8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inherited language metadata and Czech one-letter preposition wrapping.

- [#275](https://github.com/stella/folio/pull/275) [`cd5d12c`](https://github.com/stella/folio/commit/cd5d12c766ea7c55e745a3f92345e7a9ffc54b22) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Correct Garamond single-line metrics for stable OOXML pagination.

- [#320](https://github.com/stella/folio/pull/320) [`f7a409b`](https://github.com/stella/folio/commit/f7a409b63c12cbbf91b562621d632a79be559f3c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored header and footer colors at full opacity.

- [#294](https://github.com/stella/folio/pull/294) [`76c74e6`](https://github.com/stella/folio/commit/76c74e670d8b7222aee79ff6de1ff6e6832310db) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve explicit first-line paragraph indents ahead of inherited hanging indents.

- [#308](https://github.com/stella/folio/pull/308) [`f56c68c`](https://github.com/stella/folio/commit/f56c68c2b9a1f7c03186617da6da869e80c4e187) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit bold formatting on automatic list markers through parsing, layout, and painting.

- [#182](https://github.com/stella/folio/pull/182) [`d53ceb4`](https://github.com/stella/folio/commit/d53ceb49b8923780335f5cfdb61f9301f6c56c32) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent stale asynchronous overlay projections through a shared framework-neutral request gate.

- [#311](https://github.com/stella/folio/pull/311) [`482e5e7`](https://github.com/stella/folio/commit/482e5e787f226a552bbe272d0816561ba9389877) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor break-only paragraph placement and exact image-only line footprints during pagination.

- [#330](https://github.com/stella/folio/pull/330) [`c3077db`](https://github.com/stella/folio/commit/c3077db29aaf7d08c4fc806df8078a4baed228c2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve cached page-boundary reconciliation across empty carrier paragraphs.

- [#291](https://github.com/stella/folio/pull/291) [`4fb6ebd`](https://github.com/stella/folio/commit/4fb6ebdaef16fb2dd65ad32f0709069770caa15d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Split break-permitted table rows across flow regions at safe line boundaries.

- [#322](https://github.com/stella/folio/pull/322) [`7bcea45`](https://github.com/stella/folio/commit/7bcea45a7ecd00881696ba8b69d9519cc068c36a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve inherited and hanging-indent tab stops by their authored positions.

- [#300](https://github.com/stella/folio/pull/300) [`a448210`](https://github.com/stella/folio/commit/a4482101cd301d7f74202f06ed9bdc473ae0fcd9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep zero-size styled table borders visible without inflating row layout.

- [#288](https://github.com/stella/folio/pull/288) [`34a89e2`](https://github.com/stella/folio/commit/34a89e2ac88dd0467faf064ac8f7b16970e2036d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored pair kerning in worker measurement requests.

- [#302](https://github.com/stella/folio/pull/302) [`314b946`](https://github.com/stella/folio/commit/314b9460c021ba9fb9073e3cab78718648eb927f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Complete every bundled editor locale with translated UI messages.

- [#186](https://github.com/stella/folio/pull/186) [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate tracked additions when comparing against an empty document.

- [#297](https://github.com/stella/folio/pull/297) [`e11cb5b`](https://github.com/stella/folio/commit/e11cb5bd0f69d8242cab6181fd3796721eb2544b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep shallow full-hanging justified list continuations within their measured line width.

- [#295](https://github.com/stella/folio/pull/295) [`9ebad67`](https://github.com/stella/folio/commit/9ebad6730ad650fcc58fbe4f06cb70a72cc0f9bc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor the default next-page behavior for omitted section start types.

- [#284](https://github.com/stella/folio/pull/284) [`2c3992f`](https://github.com/stella/folio/commit/2c3992fbfe1a90223a01eab0d39564b78b743ed5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Respect authored pair-kerning thresholds in layout measurement and rendering.

- [#335](https://github.com/stella/folio/pull/335) [`e69d8b3`](https://github.com/stella/folio/commit/e69d8b3bebdda98b2bdb526869d5cb431e34331a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Consume redundant empty-paragraph spacing when split table rows continue in a fresh flow region.

- [#319](https://github.com/stella/folio/pull/319) [`5b2a962`](https://github.com/stella/folio/commit/5b2a962b4fe1764f0ab41aaa4233cd355fb7fdf3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored paragraph spacing after section boundaries with cached pagination hints.

- [#186](https://github.com/stella/folio/pull/186) [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve the order of consecutive trailing blocks in generated redlines.

- [#334](https://github.com/stella/folio/pull/334) [`9aeef8e`](https://github.com/stella/folio/commit/9aeef8e1650f0935830299d4f624fba34de6695a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored subpixel border weight and paragraph rule extents.

- [#292](https://github.com/stella/folio/pull/292) [`43bc9c6`](https://github.com/stella/folio/commit/43bc9c6ea4b69ffe26faba85495e515b92b2aea7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent paragraph-mark emphasis from leaking through character-style references.

- [#329](https://github.com/stella/folio/pull/329) [`9ea3bc3`](https://github.com/stella/folio/commit/9ea3bc353f4c98701b0a150958244b939a02147d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match DOCX custom kinsoku replacement lists and adjacent hanging punctuation breaks.

- [#332](https://github.com/stella/folio/pull/332) [`645a61e`](https://github.com/stella/folio/commit/645a61eba27b653e76d7db851a08e263a870ee33) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render theme-based table cell backgrounds while preserving their source theme metadata.

- [#276](https://github.com/stella/folio/pull/276) [`f5e9744`](https://github.com/stella/folio/commit/f5e97444c4e38c40ff19488dc0e828fc663925cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render grouped shape outline colours supplied by OOXML style references.

- [#310](https://github.com/stella/folio/pull/310) [`eb8f5f5`](https://github.com/stella/folio/commit/eb8f5f5bc84703fb89d73b46a3ffaf65581ca500) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Clip ordinary table-cell content to the visible row slice across page continuations.

- [#328](https://github.com/stella/folio/pull/328) [`166db7f`](https://github.com/stella/folio/commit/166db7fe854ddaac94c7739c7c64caa601313027) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve tracked insertions and deletions inside inline DOCX content controls.

- [#289](https://github.com/stella/folio/pull/289) [`f91757f`](https://github.com/stella/folio/commit/f91757f5c494b15a7f7dc734adaaf6612b69397a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve cached page boundaries on keep-with-next paragraphs.

- [#324](https://github.com/stella/folio/pull/324) [`3ec2032`](https://github.com/stella/folio/commit/3ec20321c252f664046fdeffc6da5f6509d76b32) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align table cell measurement, painting, and row splitting around one paragraph-spacing flow.

- [#321](https://github.com/stella/folio/pull/321) [`59a01c7`](https://github.com/stella/folio/commit/59a01c722a2671593fb99f08da53536c89aa37b4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve page-relative positioning and crop geometry for floating images inside tables.

- [#325](https://github.com/stella/folio/pull/325) [`d8456db`](https://github.com/stella/folio/commit/d8456db16a5b62f4d3effc97a69b082782bc3d03) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Base justified line fitting on measured compressible spaces and preserve Czech one-letter prepositions across formatted runs.

- [#326](https://github.com/stella/folio/pull/326) [`a1bf5ae`](https://github.com/stella/folio/commit/a1bf5ae038e14650ebe944cb4ce21c97d81c0170) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Suppress inherited leading spacing when a cached page boundary follows a section break.

- [#186](https://github.com/stella/folio/pull/186) [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve footnote and endnote references inside tracked insertions and deletions.

- Updated dependencies [[`689dbf5`](https://github.com/stella/folio/commit/689dbf553a028864fef280b5773eeff0fbe40d26), [`f56c68c`](https://github.com/stella/folio/commit/f56c68c2b9a1f7c03186617da6da869e80c4e187), [`482e5e7`](https://github.com/stella/folio/commit/482e5e787f226a552bbe272d0816561ba9389877), [`1ec610f`](https://github.com/stella/folio/commit/1ec610f362aab68fc55807edef88974304c22bf4), [`166db7f`](https://github.com/stella/folio/commit/166db7fe854ddaac94c7739c7c64caa601313027), [`fbc7fce`](https://github.com/stella/folio/commit/fbc7fce4c977eabace64a2756c51e42e788c5370)]:
  - @stll/docx-core@0.3.0

## 0.8.0

### Minor Changes

- [#277](https://github.com/stella/folio/pull/277) [`70573c1`](https://github.com/stella/folio/commit/70573c1c1d444fc2199308419bbb53fe13bbd415) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expand the DOCX capability catalog with core document and review features.

- [#278](https://github.com/stella/folio/pull/278) [`ce5e1df`](https://github.com/stella/folio/commit/ce5e1df63388684151841f68c9599b421fcff4da) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expand the DOCX capability catalog with document structure features.

- [#272](https://github.com/stella/folio/pull/272) [`af10f08`](https://github.com/stella/folio/commit/af10f0840565680b087fd2955ba2ab7c512e628f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add DOCX conformance detection to parsed and created packages.

- [#163](https://github.com/stella/folio/pull/163) [`3898fb3`](https://github.com/stella/folio/commit/3898fb3b7e060464272bc6311b5cb8890fcd4141) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `ensureParaIds` to the server entry: a headless, in-place `w14:paraId` backfill for `.docx` buffers (document body, headers, footers, footnotes, endnotes; table-cell and text-box paragraphs included). Deterministic, idempotent, and namespace-aware, so hosts can normalize documents once at ingest and block anchors never fall back to positional `seq-` ids.

- [#233](https://github.com/stella/folio/pull/233) [`8c1bff7`](https://github.com/stella/folio/commit/8c1bff7af7e8a80142b6a97988245b2bafc3718a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add stable document outlines, bounded section reads, scoped search, live React document navigation, and parity-safe Vue navigation stubs.

- [#267](https://github.com/stella/folio/pull/267) [`93d0c2e`](https://github.com/stella/folio/commit/93d0c2ef49a038f824969870cf7946b4c6e8c6cc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add profile-aware DOCX compatibility diagnostics.

- [#279](https://github.com/stella/folio/pull/279) [`e86b7e1`](https://github.com/stella/folio/commit/e86b7e14cd972d76d6b18cc3e4145c15721f49a1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expand the DOCX capability catalog with embedded and semantic content.

### Patch Changes

- [#257](https://github.com/stella/folio/pull/257) [`209b7dd`](https://github.com/stella/folio/commit/209b7dd924bd641c9992629123282f9046fe58ff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inherited spacing on authored blank paragraphs in page furniture.

- [#251](https://github.com/stella/folio/pull/251) [`baa3bde`](https://github.com/stella/folio/commit/baa3bdefc59650dc80893cd8d2ebc0063317a60e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inherited spacing on blank paragraphs with directly formatted paragraph marks.

- [#273](https://github.com/stella/folio/pull/273) [`cf05728`](https://github.com/stella/folio/commit/cf05728643717de020f11b581ff87a7541a71e6d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve interior table border segments when adjacent cells leave shared edges unclaimed.

- [#237](https://github.com/stella/folio/pull/237) [`f82a489`](https://github.com/stella/folio/commit/f82a489b0af3f18cdcd226b2e7b10074c5ce80b1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve shape-to-text fitting for imported OOXML text boxes.

- [#256](https://github.com/stella/folio/pull/256) [`da0ebd5`](https://github.com/stella/folio/commit/da0ebd5bf5342cfe5ba56397c3d0b31670394e5d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Measure blank hard-break lines using paragraph mark typography.

- [#245](https://github.com/stella/folio/pull/245) [`417f33d`](https://github.com/stella/folio/commit/417f33d457986be0302addfcb95472aae6a5f0e7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Drop inherited list positioning when direct formatting disables numbering.

- [#266](https://github.com/stella/folio/pull/266) [`f25eab8`](https://github.com/stella/folio/commit/f25eab8586322afa2ea6f5e2c857ce2df6a6b450) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Place wide in-flow tables below active floating exclusions when they cannot fit beside them.

- [#242](https://github.com/stella/folio/pull/242) [`58c1073`](https://github.com/stella/folio/commit/58c10730458e021fceab21640aa2ddc2627b1f87) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve positioned paragraph frames during layout and pagination.

- [#244](https://github.com/stella/folio/pull/244) [`ef9b7a6`](https://github.com/stella/folio/commit/ef9b7a6797ab288bf446265e71c87abf759e08fc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve OOXML paragraph-frame spacing, keep drop caps in normal flow, and retain side wrapping for single frames.

- [#271](https://github.com/stella/folio/pull/271) [`06adc8d`](https://github.com/stella/folio/commit/06adc8dff272e2272a87c28886f456b8c74e1bd0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve omitted table grid columns and use them when measuring row cells.

- [#248](https://github.com/stella/folio/pull/248) [`5db781a`](https://github.com/stella/folio/commit/5db781a85c7e2fae2aef3b96be6a76fa883411a7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render bounded solid VML artwork as safe SVG previews.

- [#265](https://github.com/stella/folio/pull/265) [`70ccf0e`](https://github.com/stella/folio/commit/70ccf0ec4e69dd0735a9e9e39427c816c37f475e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Move page-fitting table rows intact to the next flow region.

- [#263](https://github.com/stella/folio/pull/263) [`9cac2df`](https://github.com/stella/folio/commit/9cac2dfbd976594bd0ecdad3fa1c4e13be8eafb7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep paragraph-mark vertical alignment from changing visible text runs.

- [#246](https://github.com/stella/folio/pull/246) [`76b2a73`](https://github.com/stella/folio/commit/76b2a736f90b31bf476bb00dca5ad7bcf57abc56) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor blank even-page header and footer slots when separate odd/even page furniture is enabled.

- [#149](https://github.com/stella/folio/pull/149) [`20bf228`](https://github.com/stella/folio/commit/20bf2281523fc69bccc5c4673701178e80e43775) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Move the initial-layout font-readiness logic (`collectInitialLayoutFontFaces`, `collectInitialLayoutFontFamilies`, `documentFontsAreLoaded`, `getDocumentFontSet`, `waitForInitialLayoutFonts`) out of React's `PagedEditor.tsx` into a framework-neutral `@stll/folio-core/controller/fontReadiness` module. First slice of extracting orchestration from the React God component into the core controller; behavior is unchanged (the existing font-collection test moves to core alongside the code). No public API change.

- [#268](https://github.com/stella/folio/pull/268) [`ef6e0e7`](https://github.com/stella/folio/commit/ef6e0e793fd723b12e43a3c50c561fa3055a1385) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Constrain justified prose shrink allowance for more faithful line wrapping.

- [#235](https://github.com/stella/folio/pull/235) [`3305967`](https://github.com/stella/folio/commit/3305967b2b60cb2366388747f28612df799882b9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep body content above the in-flow extent of page footers.

- [#282](https://github.com/stella/folio/pull/282) [`48c8978`](https://github.com/stella/folio/commit/48c8978d7594aa89d91575307f983cd51deff068) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Treat stale cached pagination markers as advisory after document reflow.

- [#241](https://github.com/stella/folio/pull/241) [`becca9c`](https://github.com/stella/folio/commit/becca9c26cb4f032c18731e2c5b412461c5dd85c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve numbering-level marker alignment through parsing and layout.

- [#250](https://github.com/stella/folio/pull/250) [`820b90e`](https://github.com/stella/folio/commit/820b90ee7c9394719040ebbe8742e9dcb44d08cc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use each section's content width when measuring positioned text-box exclusion.

- [#261](https://github.com/stella/folio/pull/261) [`e6781d3`](https://github.com/stella/folio/commit/e6781d35334ab739b5a2b4d87409b99619e24f1e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Continue adjacent compatible list instances after an explicit numbering restart.

- [#193](https://github.com/stella/folio/pull/193) [`20c8dac`](https://github.com/stella/folio/commit/20c8dac1eca50d8f0e70882c4c258de25a4ebf65) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Harden the paraId lifecycle: duplicate resolution in the allocator now maps the original paragraph's position through the transaction (pasting a copy above its source no longer steals the source's id), allocation transactions are excluded from paragraph change tracking, and the hex id generators can no longer mint the reserved `00000000` value.

- [#259](https://github.com/stella/folio/pull/259) [`d2dcfd3`](https://github.com/stella/folio/commit/d2dcfd37df5f46a48f814719a41feb050fb4ad7a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve fractional section geometry and tighten justified inset-list wrapping.

- [#252](https://github.com/stella/folio/pull/252) [`e4c8718`](https://github.com/stella/folio/commit/e4c871889604546377fd14b15e2c181868171913) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Position top-and-bottom image artwork using its authored page anchors.

- [#260](https://github.com/stella/folio/pull/260) [`27d4664`](https://github.com/stella/folio/commit/27d4664ded29f74ee2fa3d7444d90f9d69d3b2b5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep paintless header and footer stories out of body margin clearance.

- [#243](https://github.com/stella/folio/pull/243) [`a75286d`](https://github.com/stella/folio/commit/a75286ddb8485bece3167b853fc0e6af88f14a06) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep style-bridged numbering streams in sequence.

- [#264](https://github.com/stella/folio/pull/264) [`b8373b8`](https://github.com/stella/folio/commit/b8373b85cb1c321f7d3b698902567d81134b44b5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render positioned legacy text boxes and keep inherited section header and footer references active across framework adapters.

- [#239](https://github.com/stella/folio/pull/239) [`c57e8a1`](https://github.com/stella/folio/commit/c57e8a1d007ab3f646e3caf96df12325cb3d3ace) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use the final section start mode when scheduling terminal OOXML section boundaries.

- [#270](https://github.com/stella/folio/pull/270) [`3971bad`](https://github.com/stella/folio/commit/3971bad5e98f35516233b8ef3382f252132f9444) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep cached pagination markers within naturally reflowed tabbed paragraph sequences.

- [#234](https://github.com/stella/folio/pull/234) [`175a090`](https://github.com/stella/folio/commit/175a0903b5a5a69e3d02d61060ca54f3c9751ffa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Apply measured header and footer clearance to section-specific body margins.

- [#240](https://github.com/stella/folio/pull/240) [`c03e665`](https://github.com/stella/folio/commit/c03e665edd1ac441f57f66b0cc8ae243ad3e6e48) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep non-rendering numbering levels from painting synthesized decimal markers.

- [#231](https://github.com/stella/folio/pull/231) [`4d271b0`](https://github.com/stella/folio/commit/4d271b0adf996b1e5f0bb07028e62b1e3f5a9e6e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored first-line list positions when deriving layout indentation.

- [#247](https://github.com/stella/folio/pull/247) [`1a979cb`](https://github.com/stella/folio/commit/1a979cb80999ac52725d42e008b8d38887d58973) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph-style boolean formatting when paragraph-mark properties differ.

- [#269](https://github.com/stella/folio/pull/269) [`91af5d7`](https://github.com/stella/folio/commit/91af5d72af7c739f52d3fa0cf8f80bd8cbcbd12d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep vertical cell insets outside minimum table-row content heights.

- [#253](https://github.com/stella/folio/pull/253) [`7895f41`](https://github.com/stella/folio/commit/7895f411079974576d84ffebc850b727089d269a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent repeated terminal empty markers from creating a body-empty page.

- [#238](https://github.com/stella/folio/pull/238) [`5ec1b97`](https://github.com/stella/folio/commit/5ec1b978b0ca23f5d451d3b07b97d12886852f5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored paragraph spacing inside text boxes.

- Updated dependencies [[`f82a489`](https://github.com/stella/folio/commit/f82a489b0af3f18cdcd226b2e7b10074c5ce80b1), [`ef9b7a6`](https://github.com/stella/folio/commit/ef9b7a6797ab288bf446265e71c87abf759e08fc), [`af10f08`](https://github.com/stella/folio/commit/af10f0840565680b087fd2955ba2ab7c512e628f), [`06adc8d`](https://github.com/stella/folio/commit/06adc8dff272e2272a87c28886f456b8c74e1bd0), [`becca9c`](https://github.com/stella/folio/commit/becca9c26cb4f032c18731e2c5b412461c5dd85c)]:
  - @stll/docx-core@0.2.0

## 0.7.0

### Minor Changes

- [#232](https://github.com/stella/folio/pull/232) [`d249b94`](https://github.com/stella/folio/commit/d249b94c0e0d5d9c3d26bef4abaad2747ad3a151) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add structured DOCX text extraction to the server API.

### Patch Changes

- [#211](https://github.com/stella/folio/pull/211) [`d34a999`](https://github.com/stella/folio/commit/d34a9999f20981d61cf9e8196f110e8b3fd27b55) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inherited spacing on empty paragraphs with direct paragraph formatting.

- [#228](https://github.com/stella/folio/pull/228) [`fb6ddd0`](https://github.com/stella/folio/commit/fb6ddd0588c6700bec8a04f173a8b64ad9cc9d07) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Suppress automatic paragraph gaps inside continuous numbered sequences.

- [#202](https://github.com/stella/folio/pull/202) [`c10c424`](https://github.com/stella/folio/commit/c10c424f10d108363187da60d641e38800bd925f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Balance short paragraph-only continuous multi-column sections across their columns.

- [#222](https://github.com/stella/folio/pull/222) [`c375fd4`](https://github.com/stella/folio/commit/c375fd40477832a9cb52baebfa7237f744eaa5ff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match reference leading empty-outline height and avoid replaying cached page breaks after a natural paragraph continuation.

- [#217](https://github.com/stella/folio/pull/217) [`3019259`](https://github.com/stella/folio/commit/30192590594b90c85c29ae655f5417a1f16fe2f5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep trailing table separators with following page content.

- [#221](https://github.com/stella/folio/pull/221) [`58879bd`](https://github.com/stella/folio/commit/58879bd0a162d99eb78fcf1aeed85d926b701e1c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Coalesce cached page markers with pages opened by keep-next chains and paragraph continuations.

- [#210](https://github.com/stella/folio/pull/210) [`a16bc87`](https://github.com/stella/folio/commit/a16bc87d9a06f861474ed8fada3391b6b15058c0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align explicitly indented table borders with reference DOCX layout.

- [#215](https://github.com/stella/folio/pull/215) [`7d09ea8`](https://github.com/stella/folio/commit/7d09ea889078646c16153de2196de16bf8dca2f6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve hanging-indent body tab positions during line wrapping.

- [#205](https://github.com/stella/folio/pull/205) [`073f521`](https://github.com/stella/folio/commit/073f5216ce51871b6fc422cb220cf02524f93435) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored zero table indentation through style resolution and layout.

- [#214](https://github.com/stella/folio/pull/214) [`2225367`](https://github.com/stella/folio/commit/22253678fa469806ba3a75c5c3df219f8d814ac7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Collapse paintless trailing spaces when positioning aligned paragraph lines.

- [#219](https://github.com/stella/folio/pull/219) [`ae636c8`](https://github.com/stella/folio/commit/ae636c8ea0fc87caed45fe680cf5ee79081c5216) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Collapse paintless ordinary spaces at soft-wrapped line starts.

- [#227](https://github.com/stella/folio/pull/227) [`b6442f5`](https://github.com/stella/folio/commit/b6442f50658c11267144980b65ac2a0f662abea3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render grouped OOXML pictures within their authored coordinate system and extent.

- [#229](https://github.com/stella/folio/pull/229) [`43ddcd4`](https://github.com/stella/folio/commit/43ddcd43a10b6ede1c66192bf0bb09538ee9bac3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep body content below the in-flow extent of default headers.

- [#224](https://github.com/stella/folio/pull/224) [`e346c20`](https://github.com/stella/folio/commit/e346c20bbac92b84399a2c8f6e17b9d09e146b78) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Exclude non-breaking spaces from the shrink capacity of justified prose.

- [#213](https://github.com/stella/folio/pull/213) [`2de1c11`](https://github.com/stella/folio/commit/2de1c1157c448fccb522f3a02de5c4e2bcf226be) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use normal justified prose wrapping when tab-stop metadata has no corresponding tab content.

- [#226](https://github.com/stella/folio/pull/226) [`621e1c4`](https://github.com/stella/folio/commit/621e1c4aafd37ed7d83dbc60e415fada4a5917af) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep justified inset list continuations within their authored body width.

- [#218](https://github.com/stella/folio/pull/218) [`20cbba3`](https://github.com/stella/folio/commit/20cbba3f8846378b4391cf565e7d17df869aff94) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve valid sub-pixel advances when positioning aligned tab content.

- [#223](https://github.com/stella/folio/pull/223) [`15f021a`](https://github.com/stella/folio/commit/15f021a1f337c75ea7434d6cda461224f9cb4ec9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allow numeric floating-table offsets to extend into page margins.

- [#203](https://github.com/stella/folio/pull/203) [`655b698`](https://github.com/stella/folio/commit/655b6983942a0afac930f84386a3824eaf5e068c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match prose justification tolerance on numbered-list continuation lines.

- [#216](https://github.com/stella/folio/pull/216) [`5ed7452`](https://github.com/stella/folio/commit/5ed7452ffc1ead26da2ecb5403b364cb389ba9ba) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent paintless terminal table anchors from creating blank pages.

- [#208](https://github.com/stella/folio/pull/208) [`2e8fe27`](https://github.com/stella/folio/commit/2e8fe27d0ad031a0d5ebd81402171a9118aaae5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor leading authored column breaks and resume continuous-section content below the tallest column.

- [#220](https://github.com/stella/folio/pull/220) [`43c9bc7`](https://github.com/stella/folio/commit/43c9bc7d251ecad3a1558e6682ca3bb57e90f224) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Avoid whitespace-only lines when preserved spaces overflow a soft wrap.

- [#209](https://github.com/stella/folio/pull/209) [`4e479e2`](https://github.com/stella/folio/commit/4e479e2d0d7c6f9f48d25ae4a4d511143e6ce834) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Avoid painting an extra blank line for empty structural section-break paragraphs.

- [#204](https://github.com/stella/folio/pull/204) [`1e8aff9`](https://github.com/stella/folio/commit/1e8aff947b9bca3c58ab997b25687e31d43bf6a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render automatic paragraph spacing using the reference 14pt gap.

## 0.6.1

### Patch Changes

- [#192](https://github.com/stella/folio/pull/192) [`ee6d2ef`](https://github.com/stella/folio/commit/ee6d2ef6f750664e662f2931c81f3f3e0e400312) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep justified zero-left hanging list lines within the authored right margin.

- [#198](https://github.com/stella/folio/pull/198) [`8712941`](https://github.com/stella/folio/commit/87129417ad86a0f9f6579410b259f9f2b4775b46) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and paint anchored text boxes and images inside table cells.

- [#191](https://github.com/stella/folio/pull/191) [`5200874`](https://github.com/stella/folio/commit/5200874a57875acfbfbda11002a2e8bb8d79b943) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve standalone column breaks and authored unequal section columns during layout.

- [#194](https://github.com/stella/folio/pull/194) [`caa171b`](https://github.com/stella/folio/commit/caa171b36ac892ef7690db69b741ea89bcf99777) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor vertical page and margin anchors for positioned header and footer text boxes.

- [#195](https://github.com/stella/folio/pull/195) [`cd53605`](https://github.com/stella/folio/commit/cd5360527e34fadd3cb7fe58b7192258e5ace0e5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align the leading text edge of unindented tables with the document content margin.

- [#199](https://github.com/stella/folio/pull/199) [`e7fbc8d`](https://github.com/stella/folio/commit/e7fbc8d46f7973abb46e5593e6db7466e214515b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve autofit table grids that exceed their preferred width.

- [#188](https://github.com/stella/folio/pull/188) [`3a2052e`](https://github.com/stella/folio/commit/3a2052e01e742f2a00a8dcfe1990abc85e679685) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Position zero-left hanging list markers in the page margin.

- Updated dependencies [[`46c6730`](https://github.com/stella/folio/commit/46c6730ebf29daccdfac64c72fcf07702709e70f)]:
  - @stll/docx-core@0.1.1

## 0.6.0

### Minor Changes

- [#183](https://github.com/stella/folio/pull/183) [`e2e8b99`](https://github.com/stella/folio/commit/e2e8b99ea804c7446dda7fbed13a758032981a39) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add typed undo handles for committed document-operation batches.

### Patch Changes

- [#178](https://github.com/stella/folio/pull/178) [`c573ddf`](https://github.com/stella/folio/commit/c573ddface0e3826324171e3a11a6c2000e13b7a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Every user-facing string in the React editor now goes through use-intl. ~173 hardcoded English JSX literals (concentrated in the dialogs: image position/properties, page setup, footnotes, watermark, tables, hyperlinks, header/footer editing) were wired to the locale catalogs — many onto existing keys the components weren't using, ~50 new keys added to en.json and synced to all locales. Visible English output is unchanged. A new `no-untranslated-jsx-literal` oxlint rule enforces this in CI so untranslatable copy cannot land again.

- [#185](https://github.com/stella/folio/pull/185) [`f67fc55`](https://github.com/stella/folio/commit/f67fc555d1db65429555232ffabde4c55b20973c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Measure vertically merged table cells against their combined row height.

- [#180](https://github.com/stella/folio/pull/180) [`0ac0260`](https://github.com/stella/folio/commit/0ac02600744994e77238b9a0a92de09656f64e3c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - i18n quality gate and terminology glossary. A new `i18n-lint` stage in `check:i18n` verifies every translated catalog string for placeholder parity with the English source, ICU MessageFormat validity, CLDR plural-category completeness per locale, dropped plurals/exact selectors, and glossary terminology (with a ratchet baseline for future debt; the current catalogs are clean). New `glossary.json` term base mimics Microsoft Word's localized terminology (LibreOffice divergences documented), covering 46 word-processing concepts across 16 locales with forbidden nonstandard variants; three existing translations were corrected to canonical Word terms (pt-BR "Recortar", tr "Açıklama").

- [#184](https://github.com/stella/folio/pull/184) [`938fe2b`](https://github.com/stella/folio/commit/938fe2b4f4070a34fbaafe04fdeb609552860aa4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render and reserve footnotes referenced from table cells.

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
