# @stll/folio-vue

## 0.11.1

### Patch Changes

- Updated dependencies [[`45af3e8`](https://github.com/stella/folio/commit/45af3e8073136efbf65d13881e8d7420d2402600), [`34b0737`](https://github.com/stella/folio/commit/34b0737e6e110d9f2aa464a4f40ad13aead5ceeb)]:
  - @stll/folio-core@0.14.0

## 0.11.0

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

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent untrusted DOCX assets from affecting the host page. Embedded fonts are
  registered under per-document scoped family names (resolved through the font
  resolver) so a document embedding a face named after a host UI family can no
  longer shadow it page-wide, and watermark dialogs validate external image
  targets against an http/https allowlist (with a defensive guard before emitting
  an external relationship) so `file:`/UNC/other-scheme targets cannot be written
  into exported documents.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Scope review actions to their intended target. AI `undoDocumentOperations`
  always undoes the body operation instead of whatever story (header/footer)
  currently has focus; sidebar accept/reject applies only the selected revision
  rather than every tracked change in the resolved range (React and Vue); the Vue
  page-setup dialog respects read-only; Vue table properties target the active
  editor view; and Vue AI-authored comments use the requested operation author.

- [#420](https://github.com/stella/folio/pull/420) [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Sanitize hyperlink and image link URLs so pasted, programmatic, or DOCX-sourced
  `javascript:`/`data:`/`file:` targets can no longer reach the live DOM or be
  opened. Hyperlink marks are sanitized on parse (`parseDOM`), on render
  (`toDOM`), and when set/inserted/edited; image `a:hlinkClick` targets are
  sanitized at parse time; the Vue popup `window.open` path now mirrors React's
  sanitizer; and aux-click on link anchors no longer bypasses the guard. Internal
  bookmark anchors (`#name`) are preserved.
- Updated dependencies [[`75842cf`](https://github.com/stella/folio/commit/75842cf60c290af3f756e7dbea7f95671fbdea4f), [`ce930f4`](https://github.com/stella/folio/commit/ce930f4ee45d2b793ef0d625fb0598ce008cb600), [`4b6e885`](https://github.com/stella/folio/commit/4b6e88531408fc9ecb82ae7c0a71e797864fa996), [`75842cf`](https://github.com/stella/folio/commit/75842cf60c290af3f756e7dbea7f95671fbdea4f), [`64f0737`](https://github.com/stella/folio/commit/64f07378ba3f460b999a8a7bba822ed0a01e37e0), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`3229068`](https://github.com/stella/folio/commit/32290689f2256e2e601f1be6701aceb5d135169f), [`a47ee19`](https://github.com/stella/folio/commit/a47ee197d1c4a5abb47efb053d7c674c71074af5), [`21274be`](https://github.com/stella/folio/commit/21274be83afaadc9d28053c87b5ea84ea619c491)]:
  - @stll/folio-core@0.13.0

## 0.10.0

### Minor Changes

- [#397](https://github.com/stella/folio/pull/397) [`0056754`](https://github.com/stella/folio/commit/0056754d22202cf27e9cd734305b10e24cc8d899) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add shared editable footnote and endnote stories with matching React and Vue surfaces.

- [#392](https://github.com/stella/folio/pull/392) [`f04bde1`](https://github.com/stella/folio/commit/f04bde1f25af45e7155c469255c900b242cd1f32) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render configurable page margin guides in both framework adapters.

- [#390](https://github.com/stella/folio/pull/390) [`9648f04`](https://github.com/stella/folio/commit/9648f0406724b1f7f4478746f2a335162bb15bfc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Load document-embedded fonts in the Vue editor and remeasure layout when their faces become ready.

- [#405](https://github.com/stella/folio/pull/405) [`d20fb9f`](https://github.com/stella/folio/commit/d20fb9fc7ae0777077e7daf5931539e63a6672cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `highlightPassage` / `clearPassageHighlight` to `DocxEditorRef`: resolve `{ blockId, text }` to a range inside the block, scroll to it, and paint a persistent translucent passage highlight, falling back to scroll-to-block with a paragraph flash when the text no longer matches. Core exports the framework-neutral `resolvePassageRange` resolver.

- [#357](https://github.com/stella/folio/pull/357) [`997f0c1`](https://github.com/stella/folio/commit/997f0c12d4418239ebd86e9b20d2058f71cbf0e1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share content-control picker state and dispatch across adapters, and add interactive Vue dropdown and date controls.

- [#387](https://github.com/stella/folio/pull/387) [`a56ce7d`](https://github.com/stella/folio/commit/a56ce7d0018948e1e78a3139a00577e3c7858d09) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Support tracked inline formatting operations across document review surfaces.

- [#372](https://github.com/stella/folio/pull/372) [`2499909`](https://github.com/stella/folio/commit/24999097e3ad8f59a4a76b27fe96ea56bf247b1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project streaming autocomplete suggestions onto Vue's painted document pages.

- [#370](https://github.com/stella/folio/pull/370) [`45f1af9`](https://github.com/stella/folio/commit/45f1af998990c28fb82d9b22c76fcd1af90e9e2b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor inherited locales and add reactive light, dark, and system color-mode providers.

- [#361](https://github.com/stella/folio/pull/361) [`1d90ce1`](https://github.com/stella/folio/commit/1d90ce1f6f4d3a2774c9e666b243254e6d835f1c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Connect the Vue table options menu to the existing table properties dialog and command.

- [#378](https://github.com/stella/folio/pull/378) [`4d49285`](https://github.com/stella/folio/commit/4d4928544738a070bcc3b2053a89a0862b8477af) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wire the Vue Insert menu to apply and clear document watermarks.

### Patch Changes

- [#402](https://github.com/stella/folio/pull/402) [`930ffc8`](https://github.com/stella/folio/commit/930ffc8dc06604bd5279405f8d20bb720dc9673f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route formatting, focus, and history commands to the active document story consistently.

- [#376](https://github.com/stella/folio/pull/376) [`354b449`](https://github.com/stella/folio/commit/354b449f7f74f182c02b69eacb6c8eaaa193f483) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route painted content-control clicks through the shared widget controller and render usable pickers in React and Vue.

- [#360](https://github.com/stella/folio/pull/360) [`908861b`](https://github.com/stella/folio/commit/908861ba38d1277d39bd50457dcdea14f0b45b91) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fire copy, cut, and paste callbacks consistently from the shared hidden editor in both adapters.

- Updated dependencies [[`ba753fb`](https://github.com/stella/folio/commit/ba753fbb07ba71370216975962002bb59c569769), [`0056754`](https://github.com/stella/folio/commit/0056754d22202cf27e9cd734305b10e24cc8d899), [`f04bde1`](https://github.com/stella/folio/commit/f04bde1f25af45e7155c469255c900b242cd1f32), [`a96f6e5`](https://github.com/stella/folio/commit/a96f6e51908e7f04955240763f1e198bdd38f374), [`49a610a`](https://github.com/stella/folio/commit/49a610a89be5e4e479557289f973111ff6ec530c), [`d4d51c6`](https://github.com/stella/folio/commit/d4d51c627f605e9d1e335402cc3556da404e4847), [`f349951`](https://github.com/stella/folio/commit/f34995146f0ee2a7838a6cc9c501e0227b9b1250), [`c478c54`](https://github.com/stella/folio/commit/c478c540eade004a1bffbc518b29191bb18ed7d9), [`7edee13`](https://github.com/stella/folio/commit/7edee13caa2b6f5fca6defdf9a980f4ea4a81f73), [`3ca97c9`](https://github.com/stella/folio/commit/3ca97c913741711e91c24cdee46017d04802bb01), [`f3d2847`](https://github.com/stella/folio/commit/f3d284783043e162fd0e2d006dbbcdfae5b0969e), [`8359b99`](https://github.com/stella/folio/commit/8359b9949eed5d56820217e18432666c5dfd2d5c), [`b39aa26`](https://github.com/stella/folio/commit/b39aa2689c98ee688a9d59219c283832ff0abecd), [`930ffc8`](https://github.com/stella/folio/commit/930ffc8dc06604bd5279405f8d20bb720dc9673f), [`4f9c04e`](https://github.com/stella/folio/commit/4f9c04ec03a94bb25cfab3810d752470ddf18e01), [`354b449`](https://github.com/stella/folio/commit/354b449f7f74f182c02b69eacb6c8eaaa193f483), [`908861b`](https://github.com/stella/folio/commit/908861ba38d1277d39bd50457dcdea14f0b45b91), [`0c7f14d`](https://github.com/stella/folio/commit/0c7f14d48b9cddfa2ba77e79880d40962d91a9e7), [`9a7fe2f`](https://github.com/stella/folio/commit/9a7fe2f7d704ab023ac852250bd96c1c124af714), [`0b73404`](https://github.com/stella/folio/commit/0b73404618886852439a57f8d9c257a0a448709c), [`ed40fac`](https://github.com/stella/folio/commit/ed40fac565fcbe264242d5e404a8cd0ad6e3f813), [`4852343`](https://github.com/stella/folio/commit/4852343a771cd78540d8fb3d79f92fa0b988cc12), [`83434c3`](https://github.com/stella/folio/commit/83434c3e4ea9bd8e75e8532c892dfe785e5ac3a3), [`0f619d4`](https://github.com/stella/folio/commit/0f619d4404d6a09d99850105fbad62dafa06bef7), [`18c8fdf`](https://github.com/stella/folio/commit/18c8fdf1cc5deb7df5d37f4467e07c115c1a0275), [`e3d7759`](https://github.com/stella/folio/commit/e3d7759dd39e06e79823e647c6b313d178f38f5d), [`0ada098`](https://github.com/stella/folio/commit/0ada0989e2781966b3296027e6c4ec8e5e45383b), [`d20fb9f`](https://github.com/stella/folio/commit/d20fb9fc7ae0777077e7daf5931539e63a6672cf), [`16fc2e1`](https://github.com/stella/folio/commit/16fc2e18506de34f3e5cde3cf70885c1e300d40d), [`01a4dc0`](https://github.com/stella/folio/commit/01a4dc083f439fa4970a5ea03cdfa7f47e212e9f), [`1092005`](https://github.com/stella/folio/commit/1092005cea1b227febdd833bd12b1ff9662b3c31), [`9836577`](https://github.com/stella/folio/commit/983657797c25746f8eeae01a26fce5cc6b2ca352), [`997f0c1`](https://github.com/stella/folio/commit/997f0c12d4418239ebd86e9b20d2058f71cbf0e1), [`460e1c3`](https://github.com/stella/folio/commit/460e1c346c025b04a326cb243d20b2172f42265a), [`43a68c9`](https://github.com/stella/folio/commit/43a68c9c3a989683f825a5756b8fd036c7549f60), [`740b282`](https://github.com/stella/folio/commit/740b2822a9e52276b324a9c14982a0e9bebea32a), [`5f888df`](https://github.com/stella/folio/commit/5f888df00a4c11b317360d2419c0b9deabdfcce3), [`5baa329`](https://github.com/stella/folio/commit/5baa3292273d72db9279f5e335aed9f620c20229), [`2e73ef5`](https://github.com/stella/folio/commit/2e73ef56f45ed7161f14447d76a63a937372642f), [`0dd5214`](https://github.com/stella/folio/commit/0dd5214f26bdc6a82a9273290b004cbf5fee43bc), [`ba753fb`](https://github.com/stella/folio/commit/ba753fbb07ba71370216975962002bb59c569769), [`f3d2847`](https://github.com/stella/folio/commit/f3d284783043e162fd0e2d006dbbcdfae5b0969e), [`a56ce7d`](https://github.com/stella/folio/commit/a56ce7d0018948e1e78a3139a00577e3c7858d09), [`f6c6312`](https://github.com/stella/folio/commit/f6c63128e3b0aceab7a473ed55b98a8a1be79ed0), [`a40ada1`](https://github.com/stella/folio/commit/a40ada166587eff96136db8d7404480451b43e0f), [`4bd9918`](https://github.com/stella/folio/commit/4bd99187c88c38e903ea95d617aea77db664c3b1), [`3cc5128`](https://github.com/stella/folio/commit/3cc512892017afa96b44a0cb609f6f3d81ceeb7b), [`3eb34b4`](https://github.com/stella/folio/commit/3eb34b41acaffd84ac85afaf144e3dc78afebb4c), [`5fc68b8`](https://github.com/stella/folio/commit/5fc68b87e4362a726428a944d232b5417ca4dda3), [`ba753fb`](https://github.com/stella/folio/commit/ba753fbb07ba71370216975962002bb59c569769), [`3ce8392`](https://github.com/stella/folio/commit/3ce83920b8d706dd0e2084967e9705b155914a0a), [`c768500`](https://github.com/stella/folio/commit/c768500d27bacc495a154d1e9d0771cc2638d669)]:
  - @stll/folio-core@0.12.0

## 0.9.0

### Minor Changes

- [#353](https://github.com/stella/folio/pull/353) [`10f395e`](https://github.com/stella/folio/commit/10f395e216668f4a6779aa652f0d0b783e838f31) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share find contracts and matching logic across adapters and add a functional Vue find/replace binding.

- [#355](https://github.com/stella/folio/pull/355) [`f2abf34`](https://github.com/stella/folio/commit/f2abf34f20e32d1094f4459ab633720e9712e439) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share persistent header and footer editing across the React and Vue adapters.

- [#354](https://github.com/stella/folio/pull/354) [`fa23331`](https://github.com/stella/folio/commit/fa233311ff79811818ce11bab7f9854097698c93) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share collaboration module loading and remote-selection painting, and complete the Vue collaboration pipeline.

- [#350](https://github.com/stella/folio/pull/350) [`860a894`](https://github.com/stella/folio/commit/860a894e4ffe94c1492a04db8aa7fc9736f097df) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add shared rendered-DOM geometry, functional Vue selection and decoration overlays, and document-independent table cell tracking.

- [#352](https://github.com/stella/folio/pull/352) [`8c73d03`](https://github.com/stella/folio/commit/8c73d0351fbe69bb56944bbd0e33751f8f6c382d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve and reveal stable document navigation targets through the Vue editor ref.

### Patch Changes

- Updated dependencies [[`10f395e`](https://github.com/stella/folio/commit/10f395e216668f4a6779aa652f0d0b783e838f31), [`f2abf34`](https://github.com/stella/folio/commit/f2abf34f20e32d1094f4459ab633720e9712e439), [`fa23331`](https://github.com/stella/folio/commit/fa233311ff79811818ce11bab7f9854097698c93), [`860a894`](https://github.com/stella/folio/commit/860a894e4ffe94c1492a04db8aa7fc9736f097df)]:
  - @stll/folio-core@0.11.0

## 0.8.1

### Patch Changes

- Updated dependencies [[`fd5717b`](https://github.com/stella/folio/commit/fd5717b3465d12dd114467106ae552e10df96699), [`5b99d1e`](https://github.com/stella/folio/commit/5b99d1ed01d0c10cecce2850dc611fb336e3f725), [`21b7e0e`](https://github.com/stella/folio/commit/21b7e0e4ba6be4303cd351195e44e0ed305de105), [`8f8b569`](https://github.com/stella/folio/commit/8f8b569f29a2c5f78001e12c3536d8aa7b4227a5), [`b77268b`](https://github.com/stella/folio/commit/b77268bb4cb4d57f2c892138ddeba97d8f16b028), [`4a4b10f`](https://github.com/stella/folio/commit/4a4b10ff5b12cdeec66374f32daed84794bf5b9c), [`d6df33d`](https://github.com/stella/folio/commit/d6df33d7b8fb536cdc4c8125dfbdd39fbd06ef49), [`b324ce3`](https://github.com/stella/folio/commit/b324ce3154899d2e311529612d7855840b64afcc), [`d2e0c4a`](https://github.com/stella/folio/commit/d2e0c4a774381f5daae028839de9f639b8de2ad6), [`692528b`](https://github.com/stella/folio/commit/692528b42a3da66e425497cb4075f3cd82609524), [`b2f3885`](https://github.com/stella/folio/commit/b2f3885619c85a29f54a74be13a7d89a6dbee46c), [`2337ae1`](https://github.com/stella/folio/commit/2337ae1b0548982775a93600260adb2f0aaa3ecc)]:
  - @stll/folio-core@0.10.0

## 0.8.0

### Minor Changes

- [#306](https://github.com/stella/folio/pull/306) [`9972201`](https://github.com/stella/folio/commit/99722013cdbf2320a0ca90e673a2716d12c91030) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add portable document style sets, the Stella Style drafting preset, sanitized DOCX style extraction, and complete fresh-document style serialization.

### Patch Changes

- [#182](https://github.com/stella/folio/pull/182) [`d53ceb4`](https://github.com/stella/folio/commit/d53ceb49b8923780335f5cfdb61f9301f6c56c32) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent stale asynchronous overlay projections through a shared framework-neutral request gate.

- Updated dependencies [[`2b6f212`](https://github.com/stella/folio/commit/2b6f212ac4aa776c6ef61b51596c4e37e3154dd0), [`8bd16d0`](https://github.com/stella/folio/commit/8bd16d09599a2bf99f070c5a0ed686db97709d04), [`79b74e3`](https://github.com/stella/folio/commit/79b74e3a37743cdc1f05af8be11b3c0c2bc3e03e), [`fc01a2d`](https://github.com/stella/folio/commit/fc01a2d1857a9d67378e2441560998c413278397), [`d3c2816`](https://github.com/stella/folio/commit/d3c2816d2bd48d04fea6abe4924d3c93d7d0104c), [`9f0ddce`](https://github.com/stella/folio/commit/9f0ddce46fb7541af2640dc42932fc789cdf726e), [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30), [`e476fda`](https://github.com/stella/folio/commit/e476fda883fbcda84b243bb144113d2bc47ef0ff), [`c9687eb`](https://github.com/stella/folio/commit/c9687ebd0972a894ea7c96fac7331fd8702315a8), [`689dbf5`](https://github.com/stella/folio/commit/689dbf553a028864fef280b5773eeff0fbe40d26), [`cd5d12c`](https://github.com/stella/folio/commit/cd5d12c766ea7c55e745a3f92345e7a9ffc54b22), [`f7a409b`](https://github.com/stella/folio/commit/f7a409b63c12cbbf91b562621d632a79be559f3c), [`76c74e6`](https://github.com/stella/folio/commit/76c74e670d8b7222aee79ff6de1ff6e6832310db), [`f56c68c`](https://github.com/stella/folio/commit/f56c68c2b9a1f7c03186617da6da869e80c4e187), [`d53ceb4`](https://github.com/stella/folio/commit/d53ceb49b8923780335f5cfdb61f9301f6c56c32), [`482e5e7`](https://github.com/stella/folio/commit/482e5e787f226a552bbe272d0816561ba9389877), [`c3077db`](https://github.com/stella/folio/commit/c3077db29aaf7d08c4fc806df8078a4baed228c2), [`4fb6ebd`](https://github.com/stella/folio/commit/4fb6ebdaef16fb2dd65ad32f0709069770caa15d), [`7bcea45`](https://github.com/stella/folio/commit/7bcea45a7ecd00881696ba8b69d9519cc068c36a), [`a448210`](https://github.com/stella/folio/commit/a4482101cd301d7f74202f06ed9bdc473ae0fcd9), [`34a89e2`](https://github.com/stella/folio/commit/34a89e2ac88dd0467faf064ac8f7b16970e2036d), [`314b946`](https://github.com/stella/folio/commit/314b9460c021ba9fb9073e3cab78718648eb927f), [`59a581a`](https://github.com/stella/folio/commit/59a581a7f530e9d1cdc9d35702ad54f1cb5bb2a6), [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30), [`e11cb5b`](https://github.com/stella/folio/commit/e11cb5bd0f69d8242cab6181fd3796721eb2544b), [`9ebad67`](https://github.com/stella/folio/commit/9ebad6730ad650fcc58fbe4f06cb70a72cc0f9bc), [`1ec610f`](https://github.com/stella/folio/commit/1ec610f362aab68fc55807edef88974304c22bf4), [`9972201`](https://github.com/stella/folio/commit/99722013cdbf2320a0ca90e673a2716d12c91030), [`2c3992f`](https://github.com/stella/folio/commit/2c3992fbfe1a90223a01eab0d39564b78b743ed5), [`e69d8b3`](https://github.com/stella/folio/commit/e69d8b3bebdda98b2bdb526869d5cb431e34331a), [`5b2a962`](https://github.com/stella/folio/commit/5b2a962b4fe1764f0ab41aaa4233cd355fb7fdf3), [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30), [`9aeef8e`](https://github.com/stella/folio/commit/9aeef8e1650f0935830299d4f624fba34de6695a), [`43bc9c6`](https://github.com/stella/folio/commit/43bc9c6ea4b69ffe26faba85495e515b92b2aea7), [`9ea3bc3`](https://github.com/stella/folio/commit/9ea3bc353f4c98701b0a150958244b939a02147d), [`645a61e`](https://github.com/stella/folio/commit/645a61eba27b653e76d7db851a08e263a870ee33), [`f5e9744`](https://github.com/stella/folio/commit/f5e97444c4e38c40ff19488dc0e828fc663925cf), [`eb8f5f5`](https://github.com/stella/folio/commit/eb8f5f5bc84703fb89d73b46a3ffaf65581ca500), [`166db7f`](https://github.com/stella/folio/commit/166db7fe854ddaac94c7739c7c64caa601313027), [`f91757f`](https://github.com/stella/folio/commit/f91757f5c494b15a7f7dc734adaaf6612b69397a), [`3ec2032`](https://github.com/stella/folio/commit/3ec20321c252f664046fdeffc6da5f6509d76b32), [`59a01c7`](https://github.com/stella/folio/commit/59a01c722a2671593fb99f08da53536c89aa37b4), [`d8456db`](https://github.com/stella/folio/commit/d8456db16a5b62f4d3effc97a69b082782bc3d03), [`a1bf5ae`](https://github.com/stella/folio/commit/a1bf5ae038e14650ebe944cb4ce21c97d81c0170), [`354ad51`](https://github.com/stella/folio/commit/354ad5151b10bece2a0ef9086b95f1d300605a30), [`fbc7fce`](https://github.com/stella/folio/commit/fbc7fce4c977eabace64a2756c51e42e788c5370)]:
  - @stll/folio-core@0.9.0

## 0.7.0

### Minor Changes

- [#233](https://github.com/stella/folio/pull/233) [`8c1bff7`](https://github.com/stella/folio/commit/8c1bff7af7e8a80142b6a97988245b2bafc3718a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add stable document outlines, bounded section reads, scoped search, live React document navigation, and parity-safe Vue navigation stubs.

### Patch Changes

- [#264](https://github.com/stella/folio/pull/264) [`b8373b8`](https://github.com/stella/folio/commit/b8373b85cb1c321f7d3b698902567d81134b44b5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render positioned legacy text boxes and keep inherited section header and footer references active across framework adapters.

- Updated dependencies [[`209b7dd`](https://github.com/stella/folio/commit/209b7dd924bd641c9992629123282f9046fe58ff), [`baa3bde`](https://github.com/stella/folio/commit/baa3bdefc59650dc80893cd8d2ebc0063317a60e), [`70573c1`](https://github.com/stella/folio/commit/70573c1c1d444fc2199308419bbb53fe13bbd415), [`cf05728`](https://github.com/stella/folio/commit/cf05728643717de020f11b581ff87a7541a71e6d), [`f82a489`](https://github.com/stella/folio/commit/f82a489b0af3f18cdcd226b2e7b10074c5ce80b1), [`da0ebd5`](https://github.com/stella/folio/commit/da0ebd5bf5342cfe5ba56397c3d0b31670394e5d), [`417f33d`](https://github.com/stella/folio/commit/417f33d457986be0302addfcb95472aae6a5f0e7), [`f25eab8`](https://github.com/stella/folio/commit/f25eab8586322afa2ea6f5e2c857ce2df6a6b450), [`58c1073`](https://github.com/stella/folio/commit/58c10730458e021fceab21640aa2ddc2627b1f87), [`ef9b7a6`](https://github.com/stella/folio/commit/ef9b7a6797ab288bf446265e71c87abf759e08fc), [`ce5e1df`](https://github.com/stella/folio/commit/ce5e1df63388684151841f68c9599b421fcff4da), [`af10f08`](https://github.com/stella/folio/commit/af10f0840565680b087fd2955ba2ab7c512e628f), [`06adc8d`](https://github.com/stella/folio/commit/06adc8dff272e2272a87c28886f456b8c74e1bd0), [`5db781a`](https://github.com/stella/folio/commit/5db781a85c7e2fae2aef3b96be6a76fa883411a7), [`70ccf0e`](https://github.com/stella/folio/commit/70ccf0ec4e69dd0735a9e9e39427c816c37f475e), [`9cac2df`](https://github.com/stella/folio/commit/9cac2dfbd976594bd0ecdad3fa1c4e13be8eafb7), [`3898fb3`](https://github.com/stella/folio/commit/3898fb3b7e060464272bc6311b5cb8890fcd4141), [`76b2a73`](https://github.com/stella/folio/commit/76b2a736f90b31bf476bb00dca5ad7bcf57abc56), [`20bf228`](https://github.com/stella/folio/commit/20bf2281523fc69bccc5c4673701178e80e43775), [`ef6e0e7`](https://github.com/stella/folio/commit/ef6e0e793fd723b12e43a3c50c561fa3055a1385), [`3305967`](https://github.com/stella/folio/commit/3305967b2b60cb2366388747f28612df799882b9), [`48c8978`](https://github.com/stella/folio/commit/48c8978d7594aa89d91575307f983cd51deff068), [`becca9c`](https://github.com/stella/folio/commit/becca9c26cb4f032c18731e2c5b412461c5dd85c), [`820b90e`](https://github.com/stella/folio/commit/820b90ee7c9394719040ebbe8742e9dcb44d08cc), [`e6781d3`](https://github.com/stella/folio/commit/e6781d35334ab739b5a2b4d87409b99619e24f1e), [`20c8dac`](https://github.com/stella/folio/commit/20c8dac1eca50d8f0e70882c4c258de25a4ebf65), [`d2dcfd3`](https://github.com/stella/folio/commit/d2dcfd37df5f46a48f814719a41feb050fb4ad7a), [`e4c8718`](https://github.com/stella/folio/commit/e4c871889604546377fd14b15e2c181868171913), [`27d4664`](https://github.com/stella/folio/commit/27d4664ded29f74ee2fa3d7444d90f9d69d3b2b5), [`a75286d`](https://github.com/stella/folio/commit/a75286ddb8485bece3167b853fc0e6af88f14a06), [`b8373b8`](https://github.com/stella/folio/commit/b8373b85cb1c321f7d3b698902567d81134b44b5), [`c57e8a1`](https://github.com/stella/folio/commit/c57e8a1d007ab3f646e3caf96df12325cb3d3ace), [`3971bad`](https://github.com/stella/folio/commit/3971bad5e98f35516233b8ef3382f252132f9444), [`8c1bff7`](https://github.com/stella/folio/commit/8c1bff7af7e8a80142b6a97988245b2bafc3718a), [`175a090`](https://github.com/stella/folio/commit/175a0903b5a5a69e3d02d61060ca54f3c9751ffa), [`c03e665`](https://github.com/stella/folio/commit/c03e665edd1ac441f57f66b0cc8ae243ad3e6e48), [`93d0c2e`](https://github.com/stella/folio/commit/93d0c2ef49a038f824969870cf7946b4c6e8c6cc), [`4d271b0`](https://github.com/stella/folio/commit/4d271b0adf996b1e5f0bb07028e62b1e3f5a9e6e), [`1a979cb`](https://github.com/stella/folio/commit/1a979cb80999ac52725d42e008b8d38887d58973), [`91af5d7`](https://github.com/stella/folio/commit/91af5d72af7c739f52d3fa0cf8f80bd8cbcbd12d), [`7895f41`](https://github.com/stella/folio/commit/7895f411079974576d84ffebc850b727089d269a), [`5ec1b97`](https://github.com/stella/folio/commit/5ec1b978b0ca23f5d451d3b07b97d12886852f5c), [`e86b7e1`](https://github.com/stella/folio/commit/e86b7e14cd972d76d6b18cc3e4145c15721f49a1)]:
  - @stll/folio-core@0.8.0

## 0.6.1

### Patch Changes

- Updated dependencies [[`d34a999`](https://github.com/stella/folio/commit/d34a9999f20981d61cf9e8196f110e8b3fd27b55), [`fb6ddd0`](https://github.com/stella/folio/commit/fb6ddd0588c6700bec8a04f173a8b64ad9cc9d07), [`c10c424`](https://github.com/stella/folio/commit/c10c424f10d108363187da60d641e38800bd925f), [`c375fd4`](https://github.com/stella/folio/commit/c375fd40477832a9cb52baebfa7237f744eaa5ff), [`3019259`](https://github.com/stella/folio/commit/30192590594b90c85c29ae655f5417a1f16fe2f5), [`58879bd`](https://github.com/stella/folio/commit/58879bd0a162d99eb78fcf1aeed85d926b701e1c), [`a16bc87`](https://github.com/stella/folio/commit/a16bc87d9a06f861474ed8fada3391b6b15058c0), [`d249b94`](https://github.com/stella/folio/commit/d249b94c0e0d5d9c3d26bef4abaad2747ad3a151), [`7d09ea8`](https://github.com/stella/folio/commit/7d09ea889078646c16153de2196de16bf8dca2f6), [`073f521`](https://github.com/stella/folio/commit/073f5216ce51871b6fc422cb220cf02524f93435), [`2225367`](https://github.com/stella/folio/commit/22253678fa469806ba3a75c5c3df219f8d814ac7), [`ae636c8`](https://github.com/stella/folio/commit/ae636c8ea0fc87caed45fe680cf5ee79081c5216), [`b6442f5`](https://github.com/stella/folio/commit/b6442f50658c11267144980b65ac2a0f662abea3), [`43ddcd4`](https://github.com/stella/folio/commit/43ddcd43a10b6ede1c66192bf0bb09538ee9bac3), [`e346c20`](https://github.com/stella/folio/commit/e346c20bbac92b84399a2c8f6e17b9d09e146b78), [`2de1c11`](https://github.com/stella/folio/commit/2de1c1157c448fccb522f3a02de5c4e2bcf226be), [`621e1c4`](https://github.com/stella/folio/commit/621e1c4aafd37ed7d83dbc60e415fada4a5917af), [`20cbba3`](https://github.com/stella/folio/commit/20cbba3f8846378b4391cf565e7d17df869aff94), [`15f021a`](https://github.com/stella/folio/commit/15f021a1f337c75ea7434d6cda461224f9cb4ec9), [`655b698`](https://github.com/stella/folio/commit/655b6983942a0afac930f84386a3824eaf5e068c), [`5ed7452`](https://github.com/stella/folio/commit/5ed7452ffc1ead26da2ecb5403b364cb389ba9ba), [`2e8fe27`](https://github.com/stella/folio/commit/2e8fe27d0ad031a0d5ebd81402171a9118aaae5c), [`43c9bc7`](https://github.com/stella/folio/commit/43c9bc7d251ecad3a1558e6682ca3bb57e90f224), [`4e479e2`](https://github.com/stella/folio/commit/4e479e2d0d7c6f9f48d25ae4a4d511143e6ce834), [`1e8aff9`](https://github.com/stella/folio/commit/1e8aff947b9bca3c58ab997b25687e31d43bf6a3)]:
  - @stll/folio-core@0.7.0

## 0.6.0

### Minor Changes

- [#187](https://github.com/stella/folio/pull/187) [`d4aa05d`](https://github.com/stella/folio/commit/d4aa05d46546c109ffd05eb7b98460491ad7a5b9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add transactional undo for live document-operation batches.

### Patch Changes

- [#183](https://github.com/stella/folio/pull/183) [`e2e8b99`](https://github.com/stella/folio/commit/e2e8b99ea804c7446dda7fbed13a758032981a39) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add typed undo handles for committed document-operation batches.

- Updated dependencies [[`e2e8b99`](https://github.com/stella/folio/commit/e2e8b99ea804c7446dda7fbed13a758032981a39), [`c573ddf`](https://github.com/stella/folio/commit/c573ddface0e3826324171e3a11a6c2000e13b7a), [`f67fc55`](https://github.com/stella/folio/commit/f67fc555d1db65429555232ffabde4c55b20973c), [`0ac0260`](https://github.com/stella/folio/commit/0ac02600744994e77238b9a0a92de09656f64e3c), [`938fe2b`](https://github.com/stella/folio/commit/938fe2b4f4070a34fbaafe04fdeb609552860aa4)]:
  - @stll/folio-core@0.6.0

## 0.5.0

### Minor Changes

- [#126](https://github.com/stella/folio/pull/126) [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Unify the reachable zoom range across adapters. A new `@stll/folio-core/utils/zoom` module exports the canonical `ZOOM_MIN` (0.25), `ZOOM_MAX` (4), and `ZOOM_STEP` (0.1); React and Vue now source their clamp from it. The Vue adapter previously clamped zoom to 0.5-2x and now matches React at 0.25-4x. Each adapter's curated toolbar zoom-level dropdown (50-200%) is unchanged and remains an intentional subset of the reachable range.

- [#131](https://github.com/stella/folio/pull/131) [`a50ac51`](https://github.com/stella/folio/commit/a50ac516ef17bf152b842603fe14bee9572bbb80) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Port `applyDocumentOperations` to the Vue `DocxEditorRef`. The Vue adapter now implements the versioned document-operation batch API at full parity with React (delegating to core's `applyFolioDocumentOperations`, mirroring the existing `applyAIEditOperations`), and re-exports `DocxEditorApplyDocumentOperationsOptions` from the package root. Previously this ref method existed only on the React adapter.

### Patch Changes

- [#166](https://github.com/stella/folio/pull/166) [`2234792`](https://github.com/stella/folio/commit/223479229acfa7f51185cf18408528a0e3df9790) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add typed affected-target receipts to document operation results.

- [#152](https://github.com/stella/folio/pull/152) [`95aca77`](https://github.com/stella/folio/commit/95aca7790f864974de8857b23e25ffcf732ead89) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add structured diagnostics to document operation results.

- [#143](https://github.com/stella/folio/pull/143) [`8f0701e`](https://github.com/stella/folio/commit/8f0701ed1f2b527ba21820f71a82c623b473d322) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Hoist the `getPageTextFromLayout` and AI-edit block-range helpers (`resolveFolioAIBlockRange`, `clampRangeToDocSize`) into framework-neutral core modules (`@stll/folio-core/paged-layout/pageText`, `@stll/folio-core/ai-edits/blockRange`). Both were previously duplicated verbatim in the React and Vue adapters; they now share one implementation (and one test suite) in core, so the two adapters can never drift on page-text extraction or block-range resolution.

- [#164](https://github.com/stella/folio/pull/164) [`f6d6a7c`](https://github.com/stella/folio/commit/f6d6a7caab18a10b354e60dcae845d77f3a7d8a0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor mirrored left and right margins on facing document pages.

- [#138](https://github.com/stella/folio/pull/138) [`502f114`](https://github.com/stella/folio/commit/502f1140ba3fedb6b15544c85c1244ca50d0c28c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Port the Insert Symbol dialog to the React adapter, reaching full parity with Vue. The symbol catalog and search now live in a shared framework-neutral `@stll/folio-core/symbols` module (six categories, `filterSymbols`); the Vue dialog is refactored onto it so both adapters share one source of truth instead of duplicating the catalog. React gains `InsertSymbolDialog` (exported from the package root), an "Insert Symbol" toolbar button (`onInsertSymbol`), and inserts the chosen character at the cursor.

- Updated dependencies [[`6d136cf`](https://github.com/stella/folio/commit/6d136cf795b9f088e726e0524f936b3a5135db85), [`33d2d04`](https://github.com/stella/folio/commit/33d2d04797f26c5391753feaddd942a84db53f96), [`b90687b`](https://github.com/stella/folio/commit/b90687b2bf6907de8cef2c7156d8691db7df45d8), [`2234792`](https://github.com/stella/folio/commit/223479229acfa7f51185cf18408528a0e3df9790), [`30b086e`](https://github.com/stella/folio/commit/30b086eb4e9779767fc94828633a0a15de8de0e4), [`7cfef90`](https://github.com/stella/folio/commit/7cfef90a1e3bc4d7324d13025eeb72bfccf7f4ee), [`50990a0`](https://github.com/stella/folio/commit/50990a0da74086f8c283bc99a68cc7db2bfaed82), [`33d2d04`](https://github.com/stella/folio/commit/33d2d04797f26c5391753feaddd942a84db53f96), [`33538c6`](https://github.com/stella/folio/commit/33538c633d1bd206996260de8f618ed40629f2c5), [`2eb96d5`](https://github.com/stella/folio/commit/2eb96d5d8d0bf3c7cfa933201c6310184f673825), [`7bd5dd8`](https://github.com/stella/folio/commit/7bd5dd8a623d5dfae28c5b43e507c4ab4c005507), [`d4a40b2`](https://github.com/stella/folio/commit/d4a40b2e28ed37390dc0a8ffbc759ba1d068eb4f), [`bf8e841`](https://github.com/stella/folio/commit/bf8e8411e6e812f7dba5ccb998960e869607507e), [`63eee6c`](https://github.com/stella/folio/commit/63eee6ca8b0188283255c8a6cb69551016b2d5dc), [`823b172`](https://github.com/stella/folio/commit/823b172cfa3fcb3b7339fcf60b6705db7b3662cb), [`58e66c3`](https://github.com/stella/folio/commit/58e66c3ec5e56b7809c8694249c6c69509e9830a), [`c864d1c`](https://github.com/stella/folio/commit/c864d1c0071cb5c7fec1fd59b68aba22def02083), [`95aca77`](https://github.com/stella/folio/commit/95aca7790f864974de8857b23e25ffcf732ead89), [`7fcb7d7`](https://github.com/stella/folio/commit/7fcb7d7242edbf61dac44aefcad3aa6990a13b71), [`8f0701e`](https://github.com/stella/folio/commit/8f0701ed1f2b527ba21820f71a82c623b473d322), [`411f5e4`](https://github.com/stella/folio/commit/411f5e433c69cb04bc3c2ea4b8d141faa791d80b), [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736), [`7bdd026`](https://github.com/stella/folio/commit/7bdd026483d29adcc1b57d5916d4e850c2e270d5), [`50b61d1`](https://github.com/stella/folio/commit/50b61d195834f00c1f07c27b3867a78803e02a91), [`f0d1cba`](https://github.com/stella/folio/commit/f0d1cba5ca7a1830aae7703e226e731fae76524b), [`750a137`](https://github.com/stella/folio/commit/750a1378c644314b7e478b593b76d41a1dbdb4cf), [`0f6f547`](https://github.com/stella/folio/commit/0f6f5472c858a1e67135ed9205b1600b62a61314), [`f7de365`](https://github.com/stella/folio/commit/f7de36526d276f4c2d5e684da3a468990c075ac4), [`f6d6a7c`](https://github.com/stella/folio/commit/f6d6a7caab18a10b354e60dcae845d77f3a7d8a0), [`3958226`](https://github.com/stella/folio/commit/3958226cf7047b9b16d506f4e18548ecf23950e3), [`1a278bf`](https://github.com/stella/folio/commit/1a278bfe33aeab6f9c4b92923223712d82617547), [`1742abe`](https://github.com/stella/folio/commit/1742abeb2f00f3215767bc3898110463844e279a), [`685ee93`](https://github.com/stella/folio/commit/685ee93fac2b5cb573c9569cd85b9c98fe3c9bca), [`8d9808c`](https://github.com/stella/folio/commit/8d9808c9fd621011b4f73a3bf9ac6e7ee3391a80), [`6997eee`](https://github.com/stella/folio/commit/6997eee34255bfa28ef3048e7c8bcb443c0ac091), [`f1cf28b`](https://github.com/stella/folio/commit/f1cf28bccb5fd5742e45ccdcd651b500c1234950), [`3130e5f`](https://github.com/stella/folio/commit/3130e5f616254e6bb127e4c410871c05f6bb386f), [`2f1a8cb`](https://github.com/stella/folio/commit/2f1a8cbf922b6ddf4fd93f105bbdc95d3ac0faa4), [`8a3888c`](https://github.com/stella/folio/commit/8a3888c71c6f68970342a6c6cde9af7baf4c1565), [`d736b57`](https://github.com/stella/folio/commit/d736b57f07953e8faf8793d3324912fa2b8e52f0), [`476fa19`](https://github.com/stella/folio/commit/476fa1918f041d3c9ffcf097b0609d1b0d37b069), [`1fc3f2f`](https://github.com/stella/folio/commit/1fc3f2f748f4330838a56af20f7a3dbf5e5d8959), [`1f74701`](https://github.com/stella/folio/commit/1f74701d3fc2e788a078f8620f006c294c6b8b4e), [`9697a47`](https://github.com/stella/folio/commit/9697a47eb1c5b93acb374e3ed53e1bdbda47b65e), [`a5d496f`](https://github.com/stella/folio/commit/a5d496f460c56eec6c7d5c3dd45f528f13f3c144), [`502f114`](https://github.com/stella/folio/commit/502f1140ba3fedb6b15544c85c1244ca50d0c28c), [`c01b5b0`](https://github.com/stella/folio/commit/c01b5b0f70ad38c9ff878f8396a418de563ff06a), [`41bc093`](https://github.com/stella/folio/commit/41bc093036e312971532f481d4320e3c63c57f51), [`174d0ce`](https://github.com/stella/folio/commit/174d0ce25de076e9a51963547bb37b059b5b21d3), [`f03fa89`](https://github.com/stella/folio/commit/f03fa8966bcee91ed03e33a1bc9f17513960b471), [`0caa2e6`](https://github.com/stella/folio/commit/0caa2e680e0fe125811215ad7ac4a38194ffa673), [`ffc12e6`](https://github.com/stella/folio/commit/ffc12e6142b03abd4ce27983019b9256719c7b2b), [`9af69ca`](https://github.com/stella/folio/commit/9af69ca2d63fab2f6795b388184f031b61a87545), [`dc188a4`](https://github.com/stella/folio/commit/dc188a42f21db9cccdd92f292a979c71d8be358b), [`4377bf0`](https://github.com/stella/folio/commit/4377bf0db4e7390c409a232852b4d20a533edc08), [`4de7540`](https://github.com/stella/folio/commit/4de75406dea5e63b96cbf1620429b198ec6f9a33), [`c7690f9`](https://github.com/stella/folio/commit/c7690f9e77b1a048d8ea6c77676436e3cd191bad), [`0ee5fba`](https://github.com/stella/folio/commit/0ee5fba1231d11f638b56c2b78c63c8479a1fed5), [`fbaed88`](https://github.com/stella/folio/commit/fbaed88a605dd796795a221c61390448e8e77736), [`317fa96`](https://github.com/stella/folio/commit/317fa96ab0b3254b9780ca621935b6c38787a204), [`1bea7a9`](https://github.com/stella/folio/commit/1bea7a9a4a4ebf8ff263735460a27c2d3ec1241b), [`9af69ca`](https://github.com/stella/folio/commit/9af69ca2d63fab2f6795b388184f031b61a87545)]:
  - @stll/folio-core@0.5.0

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
