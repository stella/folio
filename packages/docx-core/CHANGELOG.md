# @stll/docx-core

## 0.45.0

### Minor Changes

- [#3531](https://github.com/stella/folio/pull/3531) [`c4ea819`](https://github.com/stella/folio/commit/c4ea819e3b32f853801b1e30783e49fb4112e02a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - AutoFit to contents runs in the canonical session: `setTableGrid` with no widths leaves the grid unmeasured, directly or as a tracked grid change.

- [#3525](https://github.com/stella/folio/pull/3525) [`536b933`](https://github.com/stella/folio/commit/536b933f8c315dce659d1bcf97a121c3624a09da) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Paste tables, with the paragraphs around them, directly or as suggestions: the `pasteBlocks` editor intent places clipboard blocks at the caret, and the editor pastes clipboard tables through it.

- [#3529](https://github.com/stella/folio/pull/3529) [`de6d5e8`](https://github.com/stella/folio/commit/de6d5e820179b3b89e71f67a5c5a138a621d2841) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Edit watermarks in the canonical session: `watermarkOperations` (from `storyPartsOpOf`, in `controller/canonicalWatermark`) journals a watermark change as one operation, and the Vue watermark dialog applies it instead of refusing.

- [#3558](https://github.com/stella/folio/pull/3558) [`3ab1317`](https://github.com/stella/folio/commit/3ab1317358fc94f1e6a7cdffd1225a7f6498d9d0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Direct cell merges, cell splits and column edits apply beside more pending table reviews. A merge across rows with different pending insertions or deletions keeps each row's change on the words it moves; splitting a cell with a pending property change gives each new cell a record of its own; a column edit beside a partial cell insertion or deletion, or a pending span change, lands in the same column whether the reviews are accepted or rejected. Removing a row that alone continues a vertical merge ends that merge, and a resolved row's merged cell keeps its pending paragraph marks when it is handed on.

- [#3566](https://github.com/stella/folio/pull/3566) [`cfef00c`](https://github.com/stella/folio/commit/cfef00c3cb6f39ceca395c1526e43785d4b7274a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Kept range markers record the revision IDs that hold their cut so tracked changes resolve consistently across orders, including after reopening DOCX files with older review history.

- [#3528](https://github.com/stella/folio/pull/3528) [`c6d444c`](https://github.com/stella/folio/commit/c6d444c04bd1b875ee5119884f4d6ec26c6229b6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Paste footnotes, endnotes and comments copied in the editor: `pasteBlocks` takes the `annotations` its references and markers name and copies each under a fresh id where its reference lands, directly or as a suggestion.

### Patch Changes

- [#3549](https://github.com/stella/folio/pull/3549) [`c183f97`](https://github.com/stella/folio/commit/c183f97b44c6a8f7dc21be9b6f474bcd07694fc1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct replacement removes an empty hyperlink, content control or wrapper exactly when the part of its range between two range boundaries or comment anchors removes something else, as accepting the same suggestion does. Content already deleted or moved away no longer splits that part in a direct replacement.

- [#3562](https://github.com/stella/folio/pull/3562) [`68889c7`](https://github.com/stella/folio/commit/68889c777b50affb6520eb57ec5093b800006cb5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving tracked changes keeps a range marker kept from a rejected link at the seam, in every order, when a join is held on its other side: an earlier pending insertion's, a pending deletion's, or a seam either handed over. A marker whose piece a pending deletion holds follows that deletion's seam once it is accepted, an insertion emptied by accepting the deletion it held hands its join on as accepting it does, and a removed insertion's cut goes with an earlier insertion it was made in past later insertions the same rejection removes.

- [#3555](https://github.com/stella/folio/pull/3555) [`09186cf`](https://github.com/stella/folio/commit/09186cfbe0b1349453a69d3e5215df8b2874b2e0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving tracked changes places a range marker kept from a rejected link the same way in more orders: the piece it belongs to passes to the pending change that takes the seam of a removed or emptied change, a one-sided piece that is a pending insertion takes the marker, and accepting the change a marker named leaves it naming nothing.

- [#3221](https://github.com/stella/folio/pull/3221) [`23fbe42`](https://github.com/stella/folio/commit/23fbe42eb5be676fde9ff58085027063c7b8550e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Materialize eligible plain paragraph review resolutions without joins in a single splice while preserving the native operation checks and undo effects.

- [#3221](https://github.com/stella/folio/pull/3221) [`23fbe42`](https://github.com/stella/folio/commit/23fbe42eb5be676fde9ff58085027063c7b8550e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve eligible contiguous chains of rejected tracked paragraph marks in one splice while retaining the individual native checks, receivers, and undo effects.

- [#3550](https://github.com/stella/folio/pull/3550) [`45bb0d7`](https://github.com/stella/folio/commit/45bb0d7faae91fe138b0aa03def346159ed337f3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A canonical merge drops its separator only where, as the merge runs, every paragraph after the joined one is pending removal: a deletion the batch asks for but refuses, or makes only after the merge, no longer takes the separator away, so accepting a tracked or suggested merge reads as applying it directly. `@stll/docx-core/ops` exports `paragraphHoldsContent` and `paragraphHoldsContentOnceAccepted`.

- [#3525](https://github.com/stella/folio/pull/3525) [`536b933`](https://github.com/stella/folio/commit/536b933f8c315dce659d1bcf97a121c3624a09da) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested paste beside content pending removal keeps no extra paragraph for it once accepted.

- [#3221](https://github.com/stella/folio/pull/3221) [`23fbe42`](https://github.com/stella/folio/commit/23fbe42eb5be676fde9ff58085027063c7b8550e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve a contiguous run of rejected inserted paragraph marks with one splice inside a larger story.

- [#3550](https://github.com/stella/folio/pull/3550) [`45bb0d7`](https://github.com/stella/folio/commit/45bb0d7faae91fe138b0aa03def346159ed337f3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion that takes back an author's own paragraph inserted after a container's last paragraph keeps the pending property changes that author made to that last paragraph since, so accepting reads as deleting the inserted paragraph directly, and rejecting still restores the original.

- [#3564](https://github.com/stella/folio/pull/3564) [`b154b42`](https://github.com/stella/folio/commit/b154b42a9ac207c5c64d653dcbb95db655a4d447) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph IDs when resolving splits of pending insertions, including after saving and reopening a DOCX.

- [#3552](https://github.com/stella/folio/pull/3552) [`504f7a8`](https://github.com/stella/folio/commit/504f7a8887252048d05c8183bd7b75b186935629) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked or suggested paragraph put directly after one that ends a section, before a content control, preserved block or table, applies as the direct insertion: it opens the next section and holds its own inserted mark

- [#3548](https://github.com/stella/folio/pull/3548) [`b6eeb79`](https://github.com/stella/folio/commit/b6eeb79aeb000c4e5553e7ba456c53bc0835e37e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Sequencing rebases an edit over a concurrent deletion whose ends name zero-width leaves when the edit has no position at an offset the deletion spans, instead of refusing every such pair. `trackedOrderPair` leaves out the pairs the transform still refuses, through the same `zeroWidthSpanOverlap` check.

## 0.44.0

### Minor Changes

- [#2852](https://github.com/stella/folio/pull/2852) [`28dc700`](https://github.com/stella/folio/commit/28dc700c0ccb971f81458beabaea27e79e76b657) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stage comparison resources, section properties, imported stories and table geometry through canonical semantic operations, with atomic whole-package projection and exact undo.

- [#2735](https://github.com/stella/folio/pull/2735) [`b542508`](https://github.com/stella/folio/commit/b5425080f8da87c402d21d661ec0185afbdd1188) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Choosing a value in a dropdown, combo box, date or checkbox content control changes the control in the canonical session. The control's value (`w:lastValue`, `w:fullDate`, `w14:checked`) and its shown content change together as one undoable edit, suggesting mode tracks the shown content, and the value survives save and reopen. A locked control refuses the change and the editor says so; a refused pick no longer reports success. The new `setContentControl` operation sets a block-level or inline control's properties by its id, with an exact inverse, and is sequenced like other operations that address the paragraphs a record holds.

- [#3504](https://github.com/stella/folio/pull/3504) [`1d671e4`](https://github.com/stella/folio/commit/1d671e4b345d92637cb6560de094126973719277) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Support namespace prefix validation in the XML reader and use it when checking changed XML parts in CLI saves.

- [#3265](https://github.com/stella/folio/pull/3265) [`f565dd6`](https://github.com/stella/folio/commit/f565dd610f75f7e517d5d4e9387349603b77c223) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Column edits work on tables whose rows end early or hold no cells: a row that stops short omits the columns after its cells, as a stated `gridAfter` does. Deleting a column takes a row with it when the row loses its only cell (directly, or when the suggested deletion is accepted), and removing the column that holds every cell removes the table. Rejecting a suggested column insertion on such a table restores it. `DOCUMENT_OP_REFUSAL_REASONS.TABLE_ROW_EMPTY` is removed: no edit refuses for that reason any more.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Report the pending property changes an operation folds its edit into. A tracked docx-core operation lists them in the new `AppliedDocumentOp.folded`, beside the new ids in `revisions`, and `combineRecordedRevisions` combines both lists over operations applied in turn. A canonical batch receipt names the id of a run or paragraph property change its operation now shares, so accepting or rejecting the receipt's ids covers the operation's edit.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Run a batch's formatting operations ahead of its comments and in input order, so a later write of a property replaces an earlier one wherever their ranges meet, as applying them one at a time does. A batch order address of an annotation names its kind (`BATCH_ANNOTATION_KINDS`). A canonical batch reports an operation that the batch's earlier operations leave with nothing to change as `noopOperation`.

- [#3290](https://github.com/stella/folio/pull/3290) [`9521693`](https://github.com/stella/folio/commit/9521693e250d936ed9d0b18e4d82d51c6a535250) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add the `containerStart` insertion point under operation schema 17: paragraphs or a table go before the first block of a story, cell, note, text box or block content control whatever that block is, such as an empty content control, direct or tracked, with exact inverses; a tracked paragraph inserted beside a table may now stand before a content control. Older operation envelopes remain explicitly refused. On the canonical session, a paragraph or table inserted before an empty content control that starts its container goes before it, one inserted beside an empty control a table stands next to goes between them, and table operations address a table whose first cell starts with an empty content control.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Save a kept range marker's cut as folio review history (`folio:keptCut`, history version 2) and read it back, so a review saved midway resolves as the open one; version 1 history is still read.

- [#3460](https://github.com/stella/folio/pull/3460) [`56e28aa`](https://github.com/stella/folio/commit/56e28aa38b14113a998fe12aa9f65877e544df88) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A kept range marker's cut (`KeptCut`, saved as `folio:keptCut`) no longer carries `depth`: no resolution read it. The field is gone from the model, the operation wire and the saved review history; a file saved with it opens as before, and its depth is read past.

- [#3015](https://github.com/stella/folio/pull/3015) [`cd0bfd0`](https://github.com/stella/folio/commit/cd0bfd0af70432b78e58c6305bb63824151f1dc6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add listDocxPackageIssues: every schema-attribute and package-graph issue validateDocxPackage would report, not only the first.

- [#2993](https://github.com/stella/folio/pull/2993) [`6c8459c`](https://github.com/stella/folio/commit/6c8459c79bda231337fc08f488ed43a45284b2be) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accept explicit Markdown input for paragraph, list and plain table insertion operations.

- [#3458](https://github.com/stella/folio/pull/3458) [`c06e221`](https://github.com/stella/folio/commit/c06e22186fcd977e244888d7fa42685e7b4b57f0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open documents with an automatic note mark outside its matching note story, preserving the markup and reporting a `misplaced-note-mark` warning.

- [#3327](https://github.com/stella/folio/pull/3327) [`1371bfb`](https://github.com/stella/folio/commit/1371bfb812bf468179eeb8abc6d87ab66b19c5c6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Insert several adjacent columns in one `insertColumn` operation (`count`), sharing one grid change when tracked, and refuse a second tracked column insertion beside a pending grid change with its own reason, `pendingGridChange`.

- [#3437](https://github.com/stella/folio/pull/3437) [`993bda3`](https://github.com/stella/folio/commit/993bda3205afa2fe1bbec14eb567cad416a860b5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A comment owns one range. `insertContent` gains `comments` and `newCommentIds` under operation schema 18: an inserted comment anchor of a kind its comment already has standing in the document is a copy, which takes a fresh comment id and a copy of the thread the insertion carries when the slice holds the comment's range whole, and is left out otherwise; anchors completing a range, or of a comment the document does not hold, go in as they are, the slice's first range where it repeats one; anchors in the slice's text boxes follow the same rules. Fresh ids go to the copied comments in carried order, and an added thread belongs only to the insertion's own suggestion. Undoing a comment edit keeps the comment parts that comments made since need. Direct, tracked and suggested insertions follow it, with exact inverses. Older operation envelopes remain explicitly refused. On the legacy session, pasting commented text the document already holds gives the pasted copy a fresh comment id and a copy of its thread, undone with the paste, or leaves the comment out when the session holds no thread for it; a dropped copy does the same, each of several table cells a paste fills takes a copy of its own, and a move or a paste over the source itself keeps the comment on the placed text. Pasted copies over a selection keep their threads while suggesting. Plain text pasted inside a comment joins its range on both sessions.

- [#3078](https://github.com/stella/folio/pull/3078) [`b5d39a9`](https://github.com/stella/folio/commit/b5d39a9e92ef7ca5ee43124c9cdcad6e70f6f13f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Edits report which paragraph a removed one joined: `TouchedBlocks.absorbed` in docx-core results, `paragraphIdentity` on tracked operations whose paragraph joined away, and `absorbedBlocks` on a canonical reviewer's `acceptAll()` / `rejectAll()`. A join keeps the second paragraph's id, except where the joining paragraph's mark records `survivor: "first"`, which keeps the first.

- [#3309](https://github.com/stella/folio/pull/3309) [`b0ea6b3`](https://github.com/stella/folio/commit/b0ea6b3daa17b0300e046773db50fdf29339567e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a tracked insertion that would take over a break pending deletion or moving away with its own reason, `pendingBreakTakeover`.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a tracked join, a tracked deletion, a tracked row operation and a tracked terminal table insertion that would record over a paragraph mark already under review as `pendingParagraphMarkDeletion` or `pendingParagraphMarkInsertion`, by what the mark has pending, instead of `revisionConflict`; the canonical session reports them under the public reasons of the same names.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a tracked insertion after a container's final paragraph whose mark is under review with the new `pendingParagraphMarkDeletion` or `pendingParagraphMarkInsertion` refusal reason, by what the mark has pending, and a tracked join only into a table under review as `joinIntoPendingTable` (any other table it cannot run into stays `untrackable`). The canonical session maps compiler refusals through one total table: `pendingBreakTakeover` reports `pendingParagraphMarkDeletion`, each pending mark reason reports the public reason of its name, and only reasons for shapes the engine does not record yet report `unsupportedBlock` with a capability gap.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a list format change of a paragraph whose mark is pending deletion with the new docx-core refusal reason `pendingParagraphJoin` (was `revisionConflict`): `changeListFormat` decides it before compiling, in every mode, for the story's paragraphs and those of its text boxes. The canonical session reports the reason as `pendingParagraphMarkDeletion`, the skip reason both engines give a property edit of such a paragraph. The compiler groups a text box's paragraphs joined by a pending mark deletion when a position names them through the story drawing the box; docx-core exports that grouping as `editorParagraphGroupsThroughTextBoxes`.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Report a tracked operation over a paragraph mark pending insertion as the new `pendingParagraphMarkInsertion` skip reason; one over a mark pending deletion stays `pendingParagraphMarkDeletion`, so the two are never reported alike.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a table operation over a table, row, cell or grid that already carries a pending change with its own reason: docx-core's `pendingTableChange` (was `revisionConflict`), reported as the new `pendingTableChange` skip reason. The legacy engine reports it too where it refuses the same shape: a tracked deletion of a row, table or column whose rows or cells are already pending insertion or deletion.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a tracked join into a table pending insertion with its own reason: docx-core's `joinIntoPendingTable`, reported as the new `pendingTableInsertion` skip reason.

- [#2597](https://github.com/stella/folio/pull/2597) [`29c388f`](https://github.com/stella/folio/commit/29c388fa62102e3f2eade93a5b7582e3c61262ea) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A read-only model value no longer passes where the writable type is expected. Every record type of the model, inline records and index-signature records included, carries an optional `WritableRecord` field (exported from `@stll/docx-core/model`) whose `DeepReadonly` form TypeScript refuses where the writable form is expected, so handing a `ReadonlyDocument`, a style result or any part of them to a function that takes the writable type fails to compile. `DeepReadonly` hands a `Date` out as `ReadonlyDate`, a date without its setters. A writable value still passes where a read-only one is expected, and object literals need nothing extra.

  Breaking for host code that lists a model type's fields: `keyof` of a model record now includes the mark, so `Record<keyof TextFormatting, …>`, `satisfies Record<keyof …>` and `Required<ParagraphFormatting>` stop compiling. Use `RecordField<TextFormatting>` (exported from `@stll/docx-core/model` and `@stll/folio-core/types/document`), a record's own field names without the mark, in place of `keyof`, and `{ [Field in RecordField<ParagraphFormatting>]-?: ParagraphFormatting[Field] }` in place of `Required<…>`. `ExhaustiveFields` classifies a record's own fields only. Host code that passes a read-only value it got from the editor to its own function taking the writable type now fails to compile: take the read-only type when the function only reads, and copy with `writableCopy` or `cloneDocumentWithParagraphPropertySources` when it writes.

  Readers that only read now take read-only values: the color, border, shading and theme readers (`resolveColor`, `resolveShadingColor`, `resolveColorToHex`, `generateThemeTintShadeMatrix`, `resolveThemeFont`, `resolveThemeFontRef`, `getMajorFont`, `getMinorFont`, `getThemeColor`, `getThemeFonts`, `statesNoBorder`, `cssBorderStyle`, `knownThemeColor`, `themeColorSlot`), the `theme` parameters and props of core, React and Vue that only read it, the table context's `cellBorderColor`, the footnote properties dialog's inputs, the paragraph attribute `spacingExplicit` and the table attribute `_resolvedIndent`. Saving a paragraph copies its `spacingExplicit`.

- [#3268](https://github.com/stella/folio/pull/3268) [`17b8622`](https://github.com/stella/folio/commit/17b862218e4c6d081bd0cd5de4be2002bba11b15) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a paragraph break inside an inline content control (`insideInlineControl`), naming the control: it lives in one paragraph. The canonical session refuses it when typing and through document operations, in every mode.

- [#3401](https://github.com/stella/folio/pull/3401) [`f5ad1c9`](https://github.com/stella/folio/commit/f5ad1c946349ace85d1803c92e077ac15d37f598) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A revision resolution reports each tracked change its joins merged into an alike change, with the change that now holds its content (`touched.absorbedRevisions`).

- [#2959](https://github.com/stella/folio/pull/2959) [`ef2c728`](https://github.com/stella/folio/commit/ef2c728e9b8f9dd3b1fb4266b4bdfd5df7fa1d80) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep comments created concurrently with a comment restoration, such as the undo of a rejected comment, and advance operation journals to schema 16.

- [#2831](https://github.com/stella/folio/pull/2831) [`beef4c0`](https://github.com/stella/folio/commit/beef4c0d887e36bbec10f4cfdb7f16fd2c593967) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore section topology when undoing note cleanup and advance operation journals to schema 14.

- [#2842](https://github.com/stella/folio/pull/2842) [`41542d4`](https://github.com/stella/folio/commit/41542d46b5f07f8837547a338ae9017e2d5c5a6c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Permit native replacements that retire complete comment hosts, with semantic thread exhaustion before direct insertion and exact tracked rejection.

- [#3414](https://github.com/stella/folio/pull/3414) [`9d106b5`](https://github.com/stella/folio/commit/9d106b59e3bdbdc32b9b03125dbbb53ffb934f14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a run's pending property change id when a tracked patch folds the whole run into it, and report only revision ids a canonical batch leaves standing. Refuse a suggestion's fold into another pending property change with the new `pendingRunPropertyChange` and `pendingParagraphPropertyChange` refusal reasons, which the canonical session reports as the skip reasons of the same names. A tracked insertion after a container's final paragraph whose properties carry a pending change is refused as `pendingParagraphPropertyChange` too.

- [#2883](https://github.com/stella/folio/pull/2883) [`6ce79f7`](https://github.com/stella/folio/commit/6ce79f731d4b090a09410871c5010dd957979056) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add table-side paragraph insertion and exact removal inverses under operation schema 15; older operation envelopes remain explicitly refused.

- [#3333](https://github.com/stella/folio/pull/3333) [`3870c36`](https://github.com/stella/folio/commit/3870c368f2c6b435c8193f17f3f328739b8bfe65) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Text an operation adds right after a note reference takes the body run's formatting, not the reference's style (`TYPED_TEXT_FORMATTING.BODY_TEXT` for `compileEditorIntent`).

- [#3278](https://github.com/stella/folio/pull/3278) [`d8d1f5f`](https://github.com/stella/folio/commit/d8d1f5fde2110d7b523065cc3d0cbfe5900d26f9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The canonical session normalizes the list rendering of paragraphs in text boxes, as a reopen reads them. docx-core adds `textBoxContentsIn`, and `withParagraphRenderings` and `withParagraphCopies` reach paragraphs in a story's text boxes.

- [#3070](https://github.com/stella/folio/pull/3070) [`e35ed22`](https://github.com/stella/folio/commit/e35ed221fd7fbf3dcafb1f19f91605ec99750d44) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Edit text box paragraphs in the canonical session: typing, public operations, tracked and suggested edits, comments and revision review reach a text box's paragraphs, and save writes them into the box (and its alternate content fallback) instead of refusing. `locateParagraphThroughTextBoxes` and `paragraphsThroughTextBoxes` find a story's paragraphs together with those of its text boxes.

- [#3069](https://github.com/stella/folio/pull/3069) [`f0ba341`](https://github.com/stella/folio/commit/f0ba341f59694de5be0d774c5f4f8ccd36fde5b0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Edit text box paragraphs as stories of their own. An operation or editor intent whose paragraphs sit in one text box is applied in that box, named through the story that draws it; revision resolution reaches revisions inside text boxes; comments open, resolve and go in text boxes; the paragraph drawing a box is reported changed with it.

- [#3343](https://github.com/stella/folio/pull/3343) [`2214fbe`](https://github.com/stella/folio/commit/2214fbec80d6a9d0fffad8d19b60dd6385725a6a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked merge's paragraph property change belongs to the merge: accepting or rejecting the merge's paragraph-mark deletion decides the property change too, on both engines and after a save and reopen. The change keeps its own revision id and names the deletion it belongs to (`ParagraphPropertyChange.resolvesWith`, saved as `folio:resolvesWith`; review history version 2, version 1 still read). A tracked merge's `revisionId` is its paragraph-mark deletion.

- [#3267](https://github.com/stella/folio/pull/3267) [`04a2eea`](https://github.com/stella/folio/commit/04a2eea423d5ee264d855bc9665efdd8e9649007) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A part holding a reference XML refuses is not opened and does not validate: an entity no declaration defines, or a character reference to a character XML 1.0 does not allow (`&#0;`, `&#x1;`, `&#xFFFE;`, a surrogate, past U+10FFFF), with digits that are not a number (`&#12ab;`), or with a hexadecimal marker in capitals (`&#X41;`). Both engines refuse it with `MalformedXmlReferenceError` (part, offset, reference and kind), which replaces `UndeclaredXmlEntityError`; `malformedXmlReferenceError` replaces `undeclaredXmlEntityError`. The legacy engine no longer reads such a reference as nothing. The package validators and the CLI's save checks refuse the same parts. `XmlReadOptions.undeclaredEntities` and `XmlUndeclaredEntities` are removed; no read accepts these references.

### Patch Changes

- [#3438](https://github.com/stella/folio/pull/3438) [`7717816`](https://github.com/stella/folio/commit/77178167741d0c5fff0347cf17267cefb5fd859c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting an inserted paragraph break whose deletion is still pending leaves that deletion holding the split's cut, with when the split was made (`splitDate` on the paragraph mark, saved as its own ignorable review-provenance attribute, `folio:splitDate`). Accepting the deletion later rejoins the same records as accepting both at once. An accepted change at the edge of a deleted break's cut retires that cut as it retires an inserted break's, so accepting the change and the deletion together leaves the pieces apart, as accepting them one at a time does.

- [#2854](https://github.com/stella/folio/pull/2854) [`af6aa54`](https://github.com/stella/folio/commit/af6aa54464e2e2775e9e9a90830b8cc3c86edfa5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep surviving range boundaries inside the recorded source container when adjacent inserted content is accepted before their source revision is rejected.

- [#2786](https://github.com/stella/folio/pull/2786) [`874b995`](https://github.com/stella/folio/commit/874b99527e3642d07b0177488f192314952fd98d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `resolvedDocumentView` resolves every revision of every story one way, as Accept All or Reject All does. Clean Markdown and the block plain-text reader read their accepted view from it, so they keep the same paragraphs, tables and notes a reviewer keeps after accepting all.

- [#2624](https://github.com/stella/folio/pull/2624) [`aef4eff`](https://github.com/stella/folio/commit/aef4efffb942282af3f658f97d6c646ededcdf3f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The reference sequencer leases each actor disjoint blocks of paragraph, revision, content-control and comment ids and rejects a batch that introduces an id outside its actor's lease. Id allocators and the canonical session mint only from a lease when they hold one; without one they keep minting the lowest free ids.

- [#3207](https://github.com/stella/folio/pull/3207) [`2d58dff`](https://github.com/stella/folio/commit/2d58dfff8c46e4cff9783a23bc3086102eab6836) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reserve paragraph identities once when compiling an ordered batch of tracked insertions.

- [#2908](https://github.com/stella/folio/pull/2908) [`7bf02cb`](https://github.com/stella/folio/commit/7bf02cb5f3a072d29c99d71b179ec4aa74a0e774) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a document's own range boundaries in place when rejecting changes joins a seam beside them: a rejected deletion's boundary content, and a boundary after a cut whose pieces the same rejection removes, no longer move into the neighbouring record.

- [#2841](https://github.com/stella/folio/pull/2841) [`e49ff6b`](https://github.com/stella/folio/commit/e49ff6b583bf7d4196acc4ee6cfe1f269147151a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve automatic note marks and their source formatting when resolving tracked content, without restoring removed words or resolved revision wrappers.

- [#3082](https://github.com/stella/folio/pull/3082) [`4c9f1ed`](https://github.com/stella/folio/commit/4c9f1ed046f9c6bbbf496075639a3e6ec5620ccb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Place a canonical batch's edits past the comment references the same batch adds before them, as the edits land when applied one at a time.

- [#3133](https://github.com/stella/folio/pull/3133) [`5e26c01`](https://github.com/stella/folio/commit/5e26c01a5db865443dc030d38cd2d1c55f3e39a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce memory use when applying operation batches by releasing intermediate document versions.

- [#2822](https://github.com/stella/folio/pull/2822) [`fe5bf61`](https://github.com/stella/folio/commit/fe5bf6107c5f600b61949c13dd734917cc278118) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of whole paragraphs now takes paragraphs holding comment range anchors: the anchors stand outside the deletion, and accepting it hands them to the nearest paragraph that stays, as deleting the paragraphs directly does. Rejecting it restores the paragraphs as they were.

- [#2805](https://github.com/stella/folio/pull/2805) [`b6b6ed3`](https://github.com/stella/folio/commit/b6b6ed3e40c168732e77a22443bb93cad5790c0f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse concurrent direct and tracked paragraph insertions after the same anchor, and direct insertions after a pending first-survivor paragraph mark, so rejecting a tracked insertion cannot discard a direct paragraph's formatting or merge away its paragraph boundary.

- [#3147](https://github.com/stella/folio/pull/3147) [`2a19bc4`](https://github.com/stella/folio/commit/2a19bc406204c5ad92ecd8190d1c1486911feee8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound the recent document versions retained by identity and paragraph indexes during large edit batches.

- [#3001](https://github.com/stella/folio/pull/3001) [`8f18a46`](https://github.com/stella/folio/commit/8f18a460e789595ea65fcdaecb63a7f63f17b70d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting tracked changes one at a time rejoins a cut record as accepting them at once does when an accepted insertion between its pieces holds only a range boundary.

- [#3142](https://github.com/stella/folio/pull/3142) [`54577f7`](https://github.com/stella/folio/commit/54577f7af6640566e0de6f71f145f7d4b2705e33) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse occupied identity prefixes and limit editor intent allocation to the paragraphs that determine identity demand.

- [#3011](https://github.com/stella/folio/pull/3011) [`1a9139d`](https://github.com/stella/folio/commit/1a9139d6db9f0656df2139a060727116b070b287) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve comment ranges ending at table or row boundaries in canonical review.

- [#3353](https://github.com/stella/folio/pull/3353) [`8260680`](https://github.com/stella/folio/commit/8260680436f4a4e8784fc8283bb2421ceaae4348) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A canonical batch refuses an operation whose tracked change or comment would carry an author no save accepts, with `invalidResult` and the save validator's reason; docx-core exports `trackedChangeAuthorIssue` and `commentAuthorIssue`, which its validator uses.

- [#2642](https://github.com/stella/folio/pull/2642) [`fabce6b`](https://github.com/stella/folio/commit/fabce6bdda8123628bdc8263cd0a220fbc993947) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A canonical session reports every published commit through `onCommitted`: its operations, their inverse, the version it reached, its origin (`input`, `undo`, `redo` or `remote`) and an operation id, minted by the session's `newOpId` option for a local commit (a random UUID by default) and the sequenced batch's own for a remote one.

  `joinCanonicalRoom` joins a canonical session to a sequenced room over a `CanonicalRoomTransport` (base snapshot identity, head revision, submit, journal subscription, id lease grant and renewal), with `createCanonicalMemoryRoom` as the in-memory reference transport around `createSequencer` and `canonicalSnapshotId` as the snapshot identity of a DOCX. Local commits are enqueued on the optimistic client and submitted in order; sequenced batches are rebased by the client and applied through `prepareRemote`; a refused local edit is taken back and reported to the host; a reconnect resubmits the unanswered batch under the same operation id; a joining session replays the journal from revision 0 and is refused on a snapshot mismatch. `@stll/docx-core/ops` exports the `BatchAck`, `BatchReject`, `BatchSubmission` and `ClientNotice` types.

- [#2675](https://github.com/stella/folio/pull/2675) [`5e8f4fc`](https://github.com/stella/folio/commit/5e8f4fc1592ef3fe29e6076e9cf16305e441cc74) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In a canonical room, edits to headers, footers and notes travel between peers: the wire decoder refused every position outside the main story, so a peer editing a note stopped following its room. The React and Vue editors (and `useDocxEditor`) take `onCanonicalHistoryDrop`, called with how many local undo and redo steps a peer's edit made unusable, so a host that mirrors the steps in its own undo stack removes as many.

- [#2649](https://github.com/stella/folio/pull/2649) [`2f1fc80`](https://github.com/stella/folio/commit/2f1fc806b19a4c27209e26e1f9491b3c7b576c41) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Comment operations can be sent in a batch and sequenced: creations (ranges, points, replies and revision comments), content edits, resolutions, deletions and their undo inverses, with bounded validation of the comment records they carry and the author and date they were sent with. A range or point comment's anchor moves like text over concurrent text and structure edits, and concurrent edits move past its reference; concurrent comment operations commute where they use different comments or different fields of one comment, and are refused with `partConflict` naming the new `comments` part otherwise. A new comment takes its place in the comment list by id. `UNSENT_OP_TYPES` is removed: every operation kind is sendable.

- [#2641](https://github.com/stella/folio/pull/2641) [`584b582`](https://github.com/stella/folio/commit/584b582e0bdb2c7dc00ce304d27a93d1bb2bc98e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Every operation kind except comment operations can be sent in a batch: tables, numbering, sections, package resources, headers and footers, notes, paragraph and inline replacement, paragraph review and their undo inverses, with bounded validation of the model records they carry. Operations without a position transform are sequenced by footprint: they pass concurrent edits of other paragraphs, tables and package parts, and an overlap is refused with a typed reason (`blockConflict`, `tableRequiresExclusiveEdit` or `partConflict` naming one of `EXCLUSIVE_PARTS`). Text typed at a paragraph's start and a concurrent format from the start now commute.

- [#2646](https://github.com/stella/folio/pull/2646) [`5acdb98`](https://github.com/stella/folio/commit/5acdb98e70a9c77aec0b9ec4515467959772241c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A canonical session applies a public `insertTable` beside a paragraph instead of refusing it: the table is built from the cell texts as the default engine builds it and placed as a paragraph beside that anchor is, a pending join included. It goes between the paragraph and the block next to it, and after a container's last paragraph it brings the empty paragraph the container ends with. Beside a table or any block other than a paragraph it stays refused under `publicOps.tableStructure`. The editor intent that inserts a table takes `paragraphAfter: "none"` to leave out the empty paragraph it otherwise writes after the table.

- [#2697](https://github.com/stella/folio/pull/2697) [`128263d`](https://github.com/stella/folio/commit/128263d1cb17e40b452f99c6368cf10353ab988c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The canonical session applies the public `insertTableColumn`, `mergeTableCells` and `splitTableCell` operations, directly or as tracked changes. A new column's cells hold the given texts, a line break starting a new paragraph as in a new row, in both engines. Merging cells leaves out cells that hold only an empty paragraph, and splitting a vertically merged cell gives each row its own cells again, each keeping its content. A suggestion the default engine cannot record, such as a split or a column through a merged cell, applies and gives the direct edit once accepted.

- [#2706](https://github.com/stella/folio/pull/2706) [`54e586a`](https://github.com/stella/folio/commit/54e586a745c7b55562ae2199e3ed88397c0e0c3b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compile direct signature tables and table insertion beside existing tables, including nested, touching and body-edge tables. Share the published insertion-mode policy across both engines and refuse unsafe placement at pending paragraph-mark gaps before mutation. A join ending in a wholly deleted paragraph retains its stopping break for table placement. Canonical whole-table signature suggestions retain their typed capability refusal until accepting them can commit the whole table directly.

- [#2631](https://github.com/stella/folio/pull/2631) [`5e3a2da`](https://github.com/stella/folio/commit/5e3a2da0bc280c41e1836eb779d7555cf1cd8b74) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The optimistic client's `receiveBroadcast` and `receiveReject` report how its document moved (`ClientDocumentChange`): a `rebased` change, a foreign revision rebased over the pending local batches or a refused batch's inverse rebased over the ones after it, or a `replaced` change, exact for the document, when a pending batch was dropped or the rebased change would not reach the replayed document. `transformPosition` moves a position across sequenced batches by paragraph id and offset, and `applyBatch` is exported.

  A canonical session applies another actor's sequenced edit with `prepareRemote(state, change)`, through the same staging and publishing path as a local edit. The edit is not an undo step: the local undo and redo steps are rebased over it, a step that no longer rebases is dropped with the steps beneath it and reported to `onHistoryDrop`, and the selection moves by paragraph id and offset. A session holding an id lease takes the ids of a remote edit as the authority admitted them under the other actor's lease, and holds them afterwards; its own lease check applies to its own edits. A commit carries the operations it applies (`ops`).

- [#3308](https://github.com/stella/folio/pull/3308) [`cb41bec`](https://github.com/stella/folio/commit/cb41becc6ca7403b36d2872675d62a1d1b490f98) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A replaced range on the canonical session writes tabs and line feeds as tab and break elements, so a comparison or public replacement keeps them.

- [#3342](https://github.com/stella/folio/pull/3342) [`74f6757`](https://github.com/stella/folio/commit/74f6757e57f40d43a128a4ab22f30f8702d08f98) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A canonical batch draws its revision ids from the free ones once the package holds ids near the largest, so a tracked edit beside a comment with the highest id applies; a leased revision id passes over the comment ids.

- [#2648](https://github.com/stella/folio/pull/2648) [`511af77`](https://github.com/stella/folio/commit/511af776245b6f4cd86e8b7c54b0d2336e159f5d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Canonical rooms gain two transports and presence. `openCanonicalTabRoom` shares a room between the tabs of one origin: the tab holding the room's Web Lock runs the authority, every tab reaches it over a BroadcastChannel, and a tab taking over the lock rebuilds the sequencer from the durable journal (`openCanonicalIndexedDbJournal`) before it leads. `openCanonicalYjsRoom` carries a room in a Yjs document for hosts that sync one: clients append to an inbox, and the authority helper a host server runs (`serveCanonicalYjsRoom`) sequences the inbox, appends to the journal and answers each request; a client refuses a journal entry written by any other Yjs client. `restoreCanonicalRoomAuthority` rebuilds a room's sequencer from its stored records. With the `presence` option, `joinCanonicalRoom` shares each peer's selection in canonical positions over the transport's presence channel, moves the other peers' selections over every sequenced batch it applies, and gives the host each peer's selection in the session's projection. A lease call a room can no longer forward, because its transport closed or stopped following the journal, is answered with the new `unavailable` lease refusal (`ID_LEASE_REFUSAL_REASONS.UNAVAILABLE`), and a tab whose journal write fails stands for the lock again instead of waiting on a term nobody leads.

- [#2756](https://github.com/stella/folio/pull/2756) [`2a21eda`](https://github.com/stella/folio/commit/2a21eda9a18b2fbda80a81945118deb65a7d76df) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allow property edits in irregular tables while preserving their topology and review snapshots; retain canonical paragraph owners through projection replacement and leave omitted-grid placeholders unowned.

- [#3241](https://github.com/stella/folio/pull/3241) [`ce90204`](https://github.com/stella/folio/commit/ce90204970b282ad86e1a82128a7d870edf8e5cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Check removed note references together when accepting or rejecting tracked changes.

- [#3314](https://github.com/stella/folio/pull/3314) [`9d3f0bc`](https://github.com/stella/folio/commit/9d3f0bc4902d601d83837e17ea246395ad96e81d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A merge or column deletion that leaves a row nothing it starts, only vertical merges from above, now removes the row, directly and on accepting a tracked column deletion. A tracked merge records that row's deletion beside its merge records, so accepting gives the direct result and rejecting the original. A merge whose addressed paragraph it drops puts the caret in the merged cell, and a merge beside a pending row insertion or deletion outside its rows no longer waits for it.

- [#3488](https://github.com/stella/folio/pull/3488) [`10a77ce`](https://github.com/stella/folio/commit/10a77cef08677e6d191a0032fbba448e2a0676e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving tracked changes beside rejected insertions leaves the same document in any order. What was made inside a rejected insertion that joined nothing keeps no join past it, and takes only the range markers kept from its own record. A cut goes with the insertion it was made in and is read between its own pieces. Range markers kept at a seam go where the seam goes: into the record beyond the insertions it waited on, or into the one an accepted deletion's pieces were cut from. A rejected split's cut goes on past an empty paragraph split off in it later, and an insertion at the edge of a split's piece is part of that piece.

- [#2795](https://github.com/stella/folio/pull/2795) [`e03602b`](https://github.com/stella/folio/commit/e03602b8c6d6a0929e92584fde553325ee9954e5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Delete table columns with the records held by their cells, refuse dangling anchors, and preserve pending row reviews through column edits.

- [#2734](https://github.com/stella/folio/pull/2734) [`e8b41fb`](https://github.com/stella/folio/commit/e8b41fbe498c7efbc9213752bcf0e4c5bc6d0a01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A comment whose range opens or closes between blocks, beside a content control, at the start or end of a table cell or of the story, can be deleted and replied to. Its markers between blocks go with the comment and come back in their exact places on undo; a comment state's inverse carries them in a new optional `boundaries` list.

- [#3136](https://github.com/stella/folio/pull/3136) [`dde3c52`](https://github.com/stella/folio/commit/dde3c52e0cd9c9f269f8fce4082f1d6780b3eb03) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Replacing the words before several comment ends at a content control's or wrapper's end puts the new words before all of them inside the one container, as accepting the tracked replacement does, instead of between them in a second container.

- [#3324](https://github.com/stella/folio/pull/3324) [`13a3aee`](https://github.com/stella/folio/commit/13a3aee965f98c4219a704bca4fdc66fc8cd253f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allocate a fresh comment id from the lowest free id once the highest annotation id is taken, instead of refusing.

- [#2771](https://github.com/stella/folio/pull/2771) [`bdb5458`](https://github.com/stella/folio/commit/bdb54584f884b973b1dde03233e0bfd0e383aa05) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A comment whose range holds only empty runs keeps its anchors and its definition when the document is prepared for editing and saved. When a save removes a comment, its range markers and reference go from inside content controls, hyperlinks, fields, tracked and moved changes, inline wrappers and text boxes as well, so the save is not refused.

- [#2987](https://github.com/stella/folio/pull/2987) [`18882b8`](https://github.com/stella/folio/commit/18882b8a6768d9b8520b629a786a7efd8d14912d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a comment whose range closes between a table's rows or a row's cells. Editing or deleting a cell paragraph in such a table no longer drops the comment's reference, adds a second range end, or leaves a deletion that cannot be accepted or rejected.

- [#3478](https://github.com/stella/folio/pull/3478) [`33ef8f8`](https://github.com/stella/folio/commit/33ef8f805b6b34930bb7ea87e2ec6923dbc661e0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Place replacement text after all adjoining comment endings, including endings separated by a comment reference.

- [#3456](https://github.com/stella/folio/pull/3456) [`062f0c9`](https://github.com/stella/folio/commit/062f0c9ec37ad073b7b70c6087f1440936f8ef69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving a tracked replacement of a comment's words one revision at a time gives the same document in either order: a pending insertion left between the runs a retired comment's markers separated joins them when it is rejected later.

- [#3383](https://github.com/stella/folio/pull/3383) [`6929864`](https://github.com/stella/folio/commit/692986472ca8c00e0b98db80727e8d224fcc69a4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Opening a package gives a comment whose range has no `w:commentReference` its reference after the range end on both sessions (new parse warning `missing-comment-reference`). A save writes it only into a paragraph it rewrites anyway, so an unedited save keeps the document part byte-identical.

- [#2868](https://github.com/stella/folio/pull/2868) [`86f22c7`](https://github.com/stella/folio/commit/86f22c7c400500ad62376671362466c58fb77aed) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compose automatic note cleanup with the original edit inverse so undo restores comment relationships and authored note declarations exactly.

- [#2667](https://github.com/stella/folio/pull/2667) [`43b0c9f`](https://github.com/stella/folio/commit/43b0c9f3fb123b25bc96d9a924b3576c92993e03) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Concurrent typing in one paragraph is sequenced by position. An `insertContent` of a closed inline slice moves later offsets like text, at one gap the later-sequenced insertion goes after the earlier one, and ranges that start or end at the insertion point keep out of the inserted content; a format or direct deletion around it is split. `splitInline` and `joinInline` keep offsets. Sequenced `insertContent` records whether it cut an identified record (`cutIdentified`). A comment anchor moves over typing in its paragraph. Two replies to one parent now commute: a reply's markers take their place after the thread's earlier comments, siblings by id.

- [#2662](https://github.com/stella/folio/pull/2662) [`e13ac4e`](https://github.com/stella/folio/commit/e13ac4e8d8d762302ec45d87f2f4ebb9dc20ec68) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Concurrent tracked changes no longer diverge: a concurrent record cut at a tracked change's edge or inside its range (another tracked edge, an insertion point, a format's end, an inline split or join, a paragraph split), a direct deletion reaching one of its edges and not the other, and two overlapping tracked deletions (nested, equal or crossing) are refused with `unsupportedPair` in either order. Tracked changes apart from every concurrent cut are rebased as before.

- [#3156](https://github.com/stella/folio/pull/3156) [`b5bcbff`](https://github.com/stella/folio/commit/b5bcbff92b26157e7aff9ef87a67f041d3e16f70) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove the empty paragraph after a block content control when the control's last paragraph is resolved away and a rejected split joins it into the paragraph before it.

- [#3392](https://github.com/stella/folio/pull/3392) [`c566752`](https://github.com/stella/folio/commit/c56675264c21ee58297dbb12425858f278b761a2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting the added break of an emptied paragraph that ends a block content control removes the paragraph whatever stands before it, and a paragraph whose break was rejected there while its inserted words stood is removed once those words are rejected too, so the inserted paragraphs' words and breaks resolve alike in any order; the marker is saved with the words.

- [#2999](https://github.com/stella/folio/pull/2999) [`6a4ea30`](https://github.com/stella/folio/commit/6a4ea3089e507ab62e5385573d15e943825efe6a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving away the mark of the paragraph that ends a block content control joins it with an empty paragraph right after the control, which goes; a paragraph with content after the control stays.

- [#3303](https://github.com/stella/folio/pull/3303) [`e6762e8`](https://github.com/stella/folio/commit/e6762e8aaed0ecaddeb22b8a5cd0f503ecba1ed4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of exactly a content control's text is recorded inside the control and marked as holding the control's text, so accepting it leaves the control, emptied, as deleting that text directly does.

- [#2781](https://github.com/stella/folio/pull/2781) [`ac09818`](https://github.com/stella/folio/commit/ac09818e556d2da358cc5b8ef2ff8d89a243448a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting an insertion at the start or end of a cut content control together with the insertion beside the control joins the cut through the control as deep as rejecting them in turn does. A cut's reach is read past a tracked change the same resolution removes at a record's edge, to the record that then meets it.

- [#2677](https://github.com/stella/folio/pull/2677) [`431926b`](https://github.com/stella/folio/commit/431926ba368d0f1e0dee744148ab1b7e8b01eb83) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph replacement, inline replacement, paragraph review or container replacement now writes every revision its new content holds, so an undo that puts a tracked change back is refused with `partConflict` against a concurrent decision of that change instead of the two orders diverging.

- [#2804](https://github.com/stella/folio/pull/2804) [`e553c64`](https://github.com/stella/folio/commit/e553c64e7773ac21c27c656321044b9ee342abe1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A revision decision (accepting or rejecting a tracked change) passes a concurrent direct positional edit, such as typing, a deletion or a format, unchanged, rather than being refused. It is refused only when the edit cut a revision record in two and gave its second half a new id, or is a tracked edit, whose records may continue a decided revision. A reviewer's decision is no longer dropped because another pending edit of the same peer, unrelated to it, was refused.

- [#3037](https://github.com/stella/folio/pull/3037) [`262f5dc`](https://github.com/stella/folio/commit/262f5dc6e792980dbf6014d9645f441645f45a8a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Sequence a revision decision past a concurrent direct paragraph join that states both of its halves, as past a direct split: a queued decision no longer drops when an earlier rejected split is undone.

- [#2917](https://github.com/stella/folio/pull/2917) [`77c6356`](https://github.com/stella/folio/commit/77c6356b3842bac445aef756fe21ffeed7036fbc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A revision decision passes a concurrent direct paragraph split or paragraph formatting unchanged, as it passes typing or a format, rather than being refused. A pending paste is no longer dropped when another peer's acceptance arrives first.

- [#2982](https://github.com/stella/folio/pull/2982) [`e050b5f`](https://github.com/stella/folio/commit/e050b5f5e6e8777216b45352ea26b65333079b8b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a pending suggested replacement when a concurrent revision decision in another paragraph arrives, instead of dropping it.

- [#2701](https://github.com/stella/folio/pull/2701) [`fb3e600`](https://github.com/stella/folio/commit/fb3e600bd7422828bf92594d41822d9f5e54d1d6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Range boundaries a rejection keeps at a seam now go only into the two pieces of the record the removed content cut. A pending insertion beside the seam, or a piece still held in another reviewer's pending deletion or move, does not take them: they stand at the seam until its pieces meet, and the seam waiting on that insertion records the date of the revision they were kept from, saved with the document. So the boundaries land in the same place whichever order the changes are accepted or rejected in.

- [#3076](https://github.com/stella/folio/pull/3076) [`778c61e`](https://github.com/stella/folio/commit/778c61e21b44d5a7381b5e28df42b238c43cf215) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Track the deletion of a paragraph right before a table whose rows are still a pending insertion as its words and mark: accepting removes the emptied paragraph, rejecting restores it. A paragraph the same author added whole goes outright, and a tracked join that would run a paragraph into such a table is refused.

- [#2763](https://github.com/stella/folio/pull/2763) [`e2fe40f`](https://github.com/stella/folio/commit/e2fe40f7bdb28d89b7577b1c339daf33da9fc607) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Deleting paragraphs directly reports the paragraph after them as modified when it takes the range boundaries, such as bookmarks, that the deleted paragraphs held. The modified paragraphs are read off the blocks the deletion writes back, so every paragraph it changes is reported.

- [#3100](https://github.com/stella/folio/pull/3100) [`2e12ba7`](https://github.com/stella/folio/commit/2e12ba7cf93b8cc4e3a44b3a7c105c678635870e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Delete the first paragraph of a story or cell before a table under review as its words and mark, as one after another block already is.

- [#3447](https://github.com/stella/folio/pull/3447) [`1e886c9`](https://github.com/stella/folio/commit/1e886c989dfbe9b21dcf79fdfe31805442ce6fb7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of the paragraph a pending join runs into deletes the whole joined paragraph on the canonical session, the words joined into it included; a deletion of a paragraph a pending deletion already takes whole reports no operation.

- [#3060](https://github.com/stella/folio/pull/3060) [`fcbbb4f`](https://github.com/stella/folio/commit/fcbbb4fb31b97e24bf05d46fdf28e9bf1b03e4fa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove a table row outright when its own author deletes it while their insertion of it is pending and nothing in it is anyone else's suggestion, as deleting over one's own insertion does.

- [#3074](https://github.com/stella/folio/pull/3074) [`e618283`](https://github.com/stella/folio/commit/e618283a57715d2e5705fa84f3aa83d847d8e008) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Let a tracked table deletion go ahead over its author's own pending column, which goes outright as deleting over one's own insertion does, and leave the comment anchors in the deleted rows standing as a paragraph deletion does.

- [#3056](https://github.com/stella/folio/pull/3056) [`954d91d`](https://github.com/stella/folio/commit/954d91ddc31e6da410aa9db35ae8f07f9975cd34) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Track the deletion of a paragraph that holds a comment's reference: the reference stands outside the deletion as the comment's range anchors do, and a direct deletion keeps it too, so a comment that goes on over earlier words keeps its reference.

- [#3376](https://github.com/stella/folio/pull/3376) [`76c78c9`](https://github.com/stella/folio/commit/76c78c95e8a0774ab8f75f6ba737eaf978d43cde) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Deleting a paragraph followed by range boundaries that stand between blocks joins it to the next paragraph across them, directly and tracked, with the boundaries where the two meet.

- [#3492](https://github.com/stella/folio/pull/3492) [`59c6ae0`](https://github.com/stella/folio/commit/59c6ae05a66adddf7c39c521949821f16464df92) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct deletion of a paragraph with range boundaries before a block that offers no paragraph to run into leaves them in the next paragraph after that block, as accepting the tracked deletion does, instead of refusing.

- [#3486](https://github.com/stella/folio/pull/3486) [`7694e4f`](https://github.com/stella/folio/commit/7694e4f43bffe8c57f4bf8d4aa707e13d740db29) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of a paragraph inside a pending join takes the words joined into it exactly where the paragraph's own mark stays, through breaks moved away and before a table too

- [#2780](https://github.com/stella/folio/pull/2780) [`dc07b80`](https://github.com/stella/folio/commit/dc07b80570bf28b0244607d3298f8b38c3e8f94d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct deletion whose range holds a note's own reference mark keeps the mark, and now records that in its sequenced effect (`keptMark` in a new `deleteRange` effect). A concurrent edit at a position the deletion closes up to its start, such as typing just after the deleted text, lands after the kept mark, as it does when the edit applies first. A direct format inside such a deletion records the mark on the records it cuts, so the deletion rebased over the format joins the pieces the mark keeps, as one record, as in the other order.

- [#3468](https://github.com/stella/folio/pull/3468) [`b26d5c8`](https://github.com/stella/folio/commit/b26d5c8feda2a25301fd9f6c21d9eee142d5b873) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Deleting a paragraph with nothing to change or record, as an empty one alone before a bookmark or one holding only range markers whose mark is already pending deletion, is a no-op or refused instead of reported as applied

- [#3182](https://github.com/stella/folio/pull/3182) [`9458d77`](https://github.com/stella/folio/commit/9458d770bceaae5bfc5920d654eb8099032b4a83) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Materialize large document block replacements without spreading them into function arguments.

- [#3287](https://github.com/stella/folio/pull/3287) [`bad87bd`](https://github.com/stella/folio/commit/bad87bd74ce27523d63e3f439844a7abf55332b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct deletion of the paragraphs before a table or block, with no table before them, removes them as accepting their tracked deletion does; the boundaries they hold go to the paragraph before, else into the block.

- [#2783](https://github.com/stella/folio/pull/2783) [`62f6e7f`](https://github.com/stella/folio/commit/62f6e7f9bc9fd4b66817604e016145147541f331) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep paired comment and move endpoints through direct and accepted tracked deletions, preserving their source containers and exact undo while removing fully covered pairs. Distinguish reused endpoint occurrences and settle revision selection through identity handoffs before publishing boundary decisions.

- [#3000](https://github.com/stella/folio/pull/3000) [`57edc5a`](https://github.com/stella/folio/commit/57edc5a152a2dfcf35fba1ed3346a3b2f030e089) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retire a comment thread when a direct edit removes every character it covers, as a direct block deletion does, so the model no longer holds a comment its save drops.

- [#2679](https://github.com/stella/folio/pull/2679) [`03834b6`](https://github.com/stella/folio/commit/03834b6b240284e58051b47d5a1e5301ae934784) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct paragraph join beside a paragraph mark pending deletion now gives the joined paragraph the properties the editing view shows for that join, whichever paragraph it keeps. Joining a paragraph whose mark is already deleted or moved away with the paragraph after it, as a direct deleteBlock of a container's last paragraph does, now applies and leaves that paragraph's words and properties unchanged.

- [#2966](https://github.com/stella/folio/pull/2966) [`e5b8948`](https://github.com/stella/folio/commit/e5b89485c60ea11f9fbcb3814be6ada4e84a11e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct text replacement keeps an empty hyperlink, content control or wrapper that stands apart from the text it replaces, as accepting the same replacement made as a suggestion does. Deleting nothing but such a container still removes it.

- [#3410](https://github.com/stella/folio/pull/3410) [`63bb80b`](https://github.com/stella/folio/commit/63bb80ba85a8abbbccefdc9a282da07b35cd63cd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct cell merge, cell split or column insertion or deletion applies on a table with pending review records and keeps them under their ids: a pending grid change takes the column too, and pending row and cell property changes take the spans and omissions the edit gives, so accepting or rejecting them gives the edit on the accepted or rejected table. The edit refuses only a record it cannot keep exact: a merge across rows of different pending insertions or deletions, a split of a span whose cell has pending property changes, a column edit beside a pending cell insertion or deletion that takes part of a column, or a pending change that records another span or merge than the edit rewrites. Tracked edits keep their refusals.

- [#3192](https://github.com/stella/folio/pull/3192) [`776009c`](https://github.com/stella/folio/commit/776009cd93be669af2759f513c5b62690b42ac1b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share paragraph replacement validation and inverse planning between native review and block operations.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An insertion made at the edge of a piece of a cut record, joined to that piece as deep as the cut, takes the cut's join when the change that made the cut is rejected: accepting the insertion then joins the pieces, as rejecting that change after accepting the insertion does.

- [#2604](https://github.com/stella/folio/pull/2604) [`cd7cd49`](https://github.com/stella/folio/commit/cd7cd49903aa328f5f5a2d85d300f178a23232f2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a run's revisions in any order now folds a content control or link piece a rejection emptied back into its other piece, as rejecting them at once does.

- [#2941](https://github.com/stella/folio/pull/2941) [`f3fe345`](https://github.com/stella/folio/commit/f3fe3451049d7cc504756debe826051d6e6ccaf3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove a container whose own content an accepted deletion took once the insertions it still holds are rejected, as accepting the deletion after rejecting them does. The acceptance marks those insertions with private resolution provenance that saved documents keep.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting a change that empties one piece of a container a pending insertion split leaves that insertion's cut with the other piece: rejecting the insertion later keeps its range markers in that piece, as rejecting it before the acceptance does.

- [#2911](https://github.com/stella/folio/pull/2911) [`c7483c6`](https://github.com/stella/folio/commit/c7483c6af883a77770601bd47b82d5989b7fbd31) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a tracked range deletion that would hold only empty hyperlinks, content controls or wrappers, and plan none for them in a suggestion: no run carries such a record, so no file or reader showed it. A suggested removal of a paragraph holding only an empty link now records just the join, or is refused when there is nothing to record.

- [#3456](https://github.com/stella/folio/pull/3456) [`062f0c9`](https://github.com/stella/folio/commit/062f0c9ec37ad073b7b70c6087f1440936f8ef69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep an inline content control after an exact text deletion. Track its selected text separately from adjacent text when the selection starts or ends inside the control.

- [#3406](https://github.com/stella/folio/pull/3406) [`726156b`](https://github.com/stella/folio/commit/726156b2f378ae81f447f1c8518c7caaea5c0a2b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of a paragraph that holds only empty containers, and whose mark a pending deletion already took, records nothing and is refused, as for an empty paragraph, instead of recording a deletion no reviewer sees.

- [#2626](https://github.com/stella/folio/pull/2626) [`a624d1b`](https://github.com/stella/folio/commit/a624d1b7c99a889b39858f2b6b7f6777492d123f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - One shared rule, `finalBreakCarrier` in `@stll/docx-core/ops`, decides which paragraph settles the deletion of a container's final paragraph: the nearest paragraph the batch keeps, or, when the batch deletes every paragraph back to the container's start or a table, the run's first paragraph, whose properties the emptied final paragraph takes.

- [#2892](https://github.com/stella/folio/pull/2892) [`444309d`](https://github.com/stella/folio/commit/444309da1ad8ddd8735f36bdb20a8ddfb078b7cd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Include final section suggestions in revision discovery and refuse unsupported native review atomically instead of silently retaining the proposal.

- [#2924](https://github.com/stella/folio/pull/2924) [`dccd4cb`](https://github.com/stella/folio/commit/dccd4cb2d08c7b198a603d3bd1aa72f9a7f29ea0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting a deleted paragraph mark before a table that ends its story no longer refuses: an emptied paragraph goes, and words run into the table's first cell.

- [#3026](https://github.com/stella/folio/pull/3026) [`951337a`](https://github.com/stella/folio/commit/951337ae59425159f55dcda1caa1750951f5b6f1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a tracked deletion of a paragraph whose pending mark deletion gave its properties to the container's deleted last paragraph, instead of leaving those properties on a paragraph nothing of it reaches.

- [#3370](https://github.com/stella/folio/pull/3370) [`3261929`](https://github.com/stella/folio/commit/32619297b951672186faeeba318311d3ac0655b7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Fold over values for maxima and minima instead of spreading arrays into `Math.max` and `Math.min`, so large tables, lists and selections stay within the engine's argument limit.

- [#3126](https://github.com/stella/folio/pull/3126) [`095b04d`](https://github.com/stella/folio/commit/095b04dc2f061c6d08bbe42dd79700915ee8db58) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Sequencing places an insertion among the zero-width children at a concurrent format's end outside the format, as the format keeps those children out, instead of refusing the pair as ambiguous.

- [#2711](https://github.com/stella/folio/pull/2711) [`f455190`](https://github.com/stella/folio/commit/f4551901ef188faadb9f58a6005aa325192fae17) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A direct format lying inside a concurrent direct deletion now leaves the same records in either order. A sequenced direct format records the records it cut at both of its ends (`cutRecords` in its `setRunProps` effect), and a deletion rebased over it joins the pieces of each that reaches past both of its ends again with one seam join at its point. The operation schema is now version 11.

- [#3233](https://github.com/stella/folio/pull/3233) [`fb689e8`](https://github.com/stella/folio/commit/fb689e81c58738becc5bca0d5670cb34f24e96d5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prepare eligible adjacent tracked paragraph insertions with one native materialization while retaining individual validation and undo effects.

- [#3388](https://github.com/stella/folio/pull/3388) [`632d02d`](https://github.com/stella/folio/commit/632d02dcf9c22fa7b2332c88e3f7eb011149925c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Text replacing exactly a hyperlink's text keeps the link, edited directly or tracked: the new text goes in a copy of the link, and accepting the tracked replacement gives the direct edit. A range reaching past the link, or text typed at its edge, still goes beside it.

- [#3456](https://github.com/stella/folio/pull/3456) [`062f0c9`](https://github.com/stella/folio/commit/062f0c9ec37ad073b7b70c6087f1440936f8ef69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Text replacing words inside a hyperlink, at its start or end, or typed between two of its leaves, stays in the one link, edited directly or tracked, and text typed on into the author's own pending insertion there stays in its piece of the link. A tracked insertion can continue a record whose piece the pending deletion before it holds; accepting the tracked replacement gives the direct edit in either resolution order, and rejecting it gives the source back. Bookmark and comment markers at the link's edges keep the link on the new text.

- [#2544](https://github.com/stella/folio/pull/2544) [`e008d2b`](https://github.com/stella/folio/commit/e008d2bd31403f7b02a43b2ac16b544bf0cf6af8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An editor intent finds the ids it may mint by asking the package's kept identity, reserved-id and paragraph-id indexes whether each id is taken, instead of copying a census of every id and walking every block of the main story for each new version.

- [#3153](https://github.com/stella/folio/pull/3153) [`048bd29`](https://github.com/stella/folio/commit/048bd29195237ae63ba4e245d0260dad9b0de37a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use package identity differences to reserve IDs during insertion admission.

- [#3143](https://github.com/stella/folio/pull/3143) [`7f8a76b`](https://github.com/stella/folio/commit/7f8a76bea5180f3ca3e37650c6f20bd2c6c65708) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Locate reviewed paragraph marks by ID and release intermediate documents as the resolver joins paragraphs.

- [#2722](https://github.com/stella/folio/pull/2722) [`cc70547`](https://github.com/stella/folio/commit/cc7054751d17a46a943f09f19341dae1ab52b62d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The optimistic client no longer drops every pending batch after one it drops. When a pending batch does not rebase over a sequenced one, or is rejected, the batches after it rebase as if it had never been made, over its inverse first, and only a batch that depends on it is dropped with it.

- [#3247](https://github.com/stella/folio/pull/3247) [`b9eb5b3`](https://github.com/stella/folio/commit/b9eb5b38869aeda31088310e4ba8425e3bfa216b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use indexed paragraph order for named comment range endpoints in the same block list.

- [#3154](https://github.com/stella/folio/pull/3154) [`45d9569`](https://github.com/stella/folio/commit/45d9569945306ec2ea7c975c4b9b7a4430459b0d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use indexed paragraph positions to resolve insertion formatting and deletion neighbors in canonical batches.

- [#3239](https://github.com/stella/folio/pull/3239) [`55b6e86`](https://github.com/stella/folio/commit/55b6e862df53a8afcb08c478a8d3bc8ae7b1f6a6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use indexed paragraph membership when combining paragraph join receipts.

- [#3122](https://github.com/stella/folio/pull/3122) [`87932d1`](https://github.com/stella/folio/commit/87932d1078f2ff1ef9897715cfc0eef303bc5a30) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use the package paragraph ID index when allocating paragraph IDs for editor operations.

- [#3123](https://github.com/stella/folio/pull/3123) [`cbe5da1`](https://github.com/stella/folio/commit/cbe5da1d0c51514812ebaff2107eb343408294e5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse the package paragraph ID index when checking insertions, splits, and replacements for collisions.

- [#3212](https://github.com/stella/folio/pull/3212) [`8d2f1ce`](https://github.com/stella/folio/commit/8d2f1ce9441de3ad024092e3ab3a28327e14cd88) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Locate tracked replacement ranges within one paragraph without repeated story walks.

- [#2808](https://github.com/stella/folio/pull/2808) [`4fa6ab1`](https://github.com/stella/folio/commit/4fa6ab126482960ceffc7346205b9969fa6e3552) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `insertText` puts its text exactly at the gap its position names, after the zero-width leaves `zeroWidthBefore` counts and before the others, and joins the run holding the leaf before that gap, else the one holding the leaf after it, whatever kind of leaf either is. Typing after a run that holds only a note's own reference mark joins that run, and text typed after a bookmark's end or a cached page break stays after it. A sequenced position at the typing's offset moves past the text exactly when it lies after that gap.

  Preserve original text-leaf boundaries and source-owned range markers through tracked insertion rejection in any order.

- [#3185](https://github.com/stella/folio/pull/3185) [`d912d3d`](https://github.com/stella/folio/commit/d912d3dfaa7ef48f998ec50fae05842dd3d5ff92) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting an inserted cell-final paragraph mark whose paragraph is left empty removes the paragraph and lets the paragraph before it end the cell, instead of leaving an extra empty paragraph in the cell.

- [#2937](https://github.com/stella/folio/pull/2937) [`a4c5dbe`](https://github.com/stella/folio/commit/a4c5dbe31d1483ee32bfd3ee67bb2a26e3d9b9ff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A link inserted over a selection is a paste of it there in both editing and suggesting modes. Inserted directly over several paragraphs, the link now lands in a paragraph that keeps the first paragraph's properties, as accepting the suggestion already did.

- [#2854](https://github.com/stella/folio/pull/2854) [`af6aa54`](https://github.com/stella/folio/commit/af6aa54464e2e2775e9e9a90830b8cc3c86edfa5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore inserted range endpoints exactly when undoing content insertion.

- [#3473](https://github.com/stella/folio/pull/3473) [`170daf0`](https://github.com/stella/folio/commit/170daf0c33223ef8410e60899f8b12e33e6cb096) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph goes directly after one that ends a section, before a content control or preserved block, into the section after it, so removing one there undoes

- [#3280](https://github.com/stella/folio/pull/3280) [`4635220`](https://github.com/stella/folio/commit/4635220e192d8bfad9f205a26df7a2c2aea65409) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked paragraph put after one that stands before a table whose mark it takes over now takes that paragraph's pending property change and inserted break with the mark, as a tracked split's second half does, instead of being refused. Accepting reads as the same insertion made directly; rejecting the insertion gives the paragraph back with its records. A break pending deletion or moving away, and a new final paragraph of its container, stay refused.

- [#3436](https://github.com/stella/folio/pull/3436) [`8f1d3b9`](https://github.com/stella/folio/commit/8f1d3b90081a2705f5a087785725ce4d4d373da7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read the item before an index without wrapping: at a list's start there is no previous item, where `.at(index - 1)` read the list's last one.

- [#3017](https://github.com/stella/folio/pull/3017) [`8c2c711`](https://github.com/stella/folio/commit/8c2c711611d8969a14a1a541802e0f91db5c3248) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving a paragraph mark with range boundaries standing between the paragraph and the next one joins the two, the boundaries where they meet, instead of refusing.

- [#3389](https://github.com/stella/folio/pull/3389) [`ae8b768`](https://github.com/stella/folio/commit/ae8b7680f68b6f9812abd3b901abe699846e61eb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Two paragraphs with bookmark or other range boundaries standing between them in the block list can be joined, directly or tracked, and the canonical engine merges them with `mergeBlockWithNext` as the default engine does. Directly, the boundaries go into the joined paragraph where the two meet; tracked, they stay between the blocks and accepting the join gives the direct result. A paragraph that ends a section still needs the next one to touch it. `@stll/docx-core/ops` exports `isBlockBoundary`. A block-list rewrite keeps each range boundary's own record, also beside an equal one.

- [#2913](https://github.com/stella/folio/pull/2913) [`660b8be`](https://github.com/stella/folio/commit/660b8beffe35e37f66d7aedc99bea6e4229bcdbf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop a paragraph join's kept boundaries at a pending deletion or move source nested inside the pieces it joins, so they stand between that change's pieces in every rejection order.

- [#2913](https://github.com/stella/folio/pull/2913) [`660b8be`](https://github.com/stella/folio/commit/660b8beffe35e37f66d7aedc99bea6e4229bcdbf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Leave a range boundary kept at a rejected split's seam between the pieces of a pending deletion or move source there, as rejecting the content after the split does, instead of joining those pieces around it.

- [#2650](https://github.com/stella/folio/pull/2650) [`0874c66`](https://github.com/stella/folio/commit/0874c6600a41d6901de48b0282e9ec675495b084) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a paragraph's inserted mark together with the pieces of a change it split no longer moves a bookmark or range boundary kept between them into a content control that follows: it stays where the rejected pieces were, as when each is rejected in turn.

- [#3072](https://github.com/stella/folio/pull/3072) [`ca97e49`](https://github.com/stella/folio/commit/ca97e49474465be3e8de674b119eb765b8d44914) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Let a suggested join of a paragraph whose mark a pending deletion already took apply: the two already read as one paragraph, so the suggestion keeps that deletion and records only the properties the join gives, instead of refusing it.

- [#3453](https://github.com/stella/folio/pull/3453) [`f2b817b`](https://github.com/stella/folio/commit/f2b817baad1fee0d2fa15819fede473d724d67d0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Joining two paragraphs gives the joined paragraph the properties the first one states, not the indentation its list level or style supplies: list items of different levels join without a direct indentation neither of them stated.

- [#3013](https://github.com/stella/folio/pull/3013) [`fcc091e`](https://github.com/stella/folio/commit/fcc091ef96a184bdb2d5ac7fa2d9d41fe9146208) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep retained bookmark endpoints at the same depth across mixed review orders.

- [#3459](https://github.com/stella/folio/pull/3459) [`f69cf13`](https://github.com/stella/folio/commit/f69cf13a81d61131deda96046461a4a6ff2654b2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph properties when a tracked deletion reaches a pending container end followed by a bookmark.

- [#3485](https://github.com/stella/folio/pull/3485) [`2c19195`](https://github.com/stella/folio/commit/2c191956ddfeeaa85ddad207a0d011f90ced1839) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve the direct edit's run formatting in tracked word replacements.

- [#2979](https://github.com/stella/folio/pull/2979) [`39d9d4e`](https://github.com/stella/folio/commit/39d9d4e0c82e7e80ea5ec8a4295138f992112752) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting tracked changes places kept boundaries the same way in every resolution order: a boundary at the far edge of a removed piece goes into the cut control's piece it touches, and a boundary kept from a removed insertion stays out of an older deletion when the cut around it rejoins.

- [#2684](https://github.com/stella/folio/pull/2684) [`becffbd`](https://github.com/stella/folio/commit/becffbd6e33a82f23276a78195fefd7d1f52f233) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A boundary kept from an insertion that cut a wrapper inside another reviewer's deletion now lands inside that wrapper, which joins again, when the insertion and the deletion are rejected together, as it does when the deletion is rejected first.

- [#2940](https://github.com/stella/folio/pull/2940) [`c8bb1b8`](https://github.com/stella/folio/commit/c8bb1b89f5d823443b30c8e16b4b2bde086bea14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove a paragraph that goes whole even when no paragraph stands before it: the range boundaries and comment anchors it holds move to the start of the first paragraph after it in the same list, instead of keeping the paragraph to carry them.

- [#2960](https://github.com/stella/folio/pull/2960) [`5a5bdc0`](https://github.com/stella/folio/commit/5a5bdc03b66f75ab19bc9c94b8177025f4f33e63) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting tracked changes places the boundaries they keep the same way whatever order the changes are resolved in: a boundary at a paragraph's edge no longer stops a paragraph join from rejoining the record beyond it, a one-sided cut keeps outside its piece any boundary that stood beside the removed record, and boundaries kept beside a cut control rejoin with it.

- [#2910](https://github.com/stella/folio/pull/2910) [`81979e2`](https://github.com/stella/folio/commit/81979e2bb59895c9676cb595f0c696691aeb7eaa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Place a range boundary kept beside a later pending insertion into the cut record it touches once that insertion is accepted, as rejecting after accepting does, so mixed accept and reject orders agree.

- [#2623](https://github.com/stella/folio/pull/2623) [`6f2d1b4`](https://github.com/stella/folio/commit/6f2d1b437f2a77906d29e401ee6c0402cc891a95) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A bookmark or move range boundary kept by a rejection stays when a pending insertion holds a copy of it, so rejecting the original and the copy gives the same result in either order.

- [#2645](https://github.com/stella/folio/pull/2645) [`5d41f67`](https://github.com/stella/folio/commit/5d41f67be6f385447ebf57fc4576e0350ecf780a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a change keeps every bookmark and move range boundary its content held unless the boundary is a copy that change owns, whatever else the story holds, and accepting a copy keeps it as inserting it directly does. So rejecting a boundary's insertion and accepting a copy of it leave the same boundaries in either order, in the engine and in the editor. A tracked insertion now records as copies a boundary its own content holds twice and both halves of a range whose other half the story holds, so rejecting it gives the paragraph back. Comment boundaries, which record no copies, still go where the story holds them otherwise.

- [#2588](https://github.com/stella/folio/pull/2588) [`989be62`](https://github.com/stella/folio/commit/989be62dcc841228cbc41cac45030b42f0701103) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A range boundary kept from rejected content goes as deep as the containers that content split, read on both sides of it, and a seam a pending deletion stands at waits on that deletion, so rejecting revisions at once or in any order gives the same document.

- [#3044](https://github.com/stella/folio/pull/3044) [`b2a3a9b`](https://github.com/stella/folio/commit/b2a3a9bfb731c74eb5c7a25beb65ec7642a2861c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting tracked changes at once places a boundary they kept as deep as rejecting them one after another does: one inserted inside a rejected insertion goes into the record the insertion's cut rejoins, and one kept beside a removal's seam goes into the record a container's inner cut rejoins there.

- [#2823](https://github.com/stella/folio/pull/2823) [`71a1207`](https://github.com/stella/folio/commit/71a1207bf9082a400d9df573d4f857a1b085f390) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a link while text inserted inside it is pending, and then accepting that text, places the link's kept range boundary inside the piece of the link the text keeps, as accepting the text first does, whichever way other insertions beside it are resolved. A range boundary another rejection kept beside it stays where it stands. Saved review provenance keeps how many boundaries a seam kept.

- [#2736](https://github.com/stella/folio/pull/2736) [`005b3ec`](https://github.com/stella/folio/commit/005b3ec1466b510652d31e28bf1252cfe3d399c5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Range boundaries a rejection keeps at a seam go again into the record their source owner names, also where the record beyond the seam is unlike it. They stand at the seam only beside a pending insertion made after the revision they were kept from, or while the other side's piece is still held in a pending deletion or move source; a seam between two such pending pieces now waits on both.

- [#2689](https://github.com/stella/folio/pull/2689) [`235ec0c`](https://github.com/stella/folio/commit/235ec0c960afba6a6fdc17ec8e1b9a61b016bf01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A boundary kept from a rejected insertion made after another reviewer's pending deletion, between its pieces, now stays between them instead of going inside the deletion it was never in, so accepting the deletion afterwards keeps it, as accepting the deletion first and rejecting the insertion afterwards does.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Kept markers beside a deferred seam stay beside the insertion accepted there unless it continues the piece past them, whichever order the batches go in.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving accepted and rejected revisions in either order places kept range markers alike: a pending insertion at the edge of an accepted insertion's content keeps its cut inside that content, and text inserted inside a rejected link takes the link's kept markers into its piece when accepted later.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A range marker kept by a resolution records the same cut whichever order resolved the revisions around it: kept cuts in a paragraph are settled once its joins stand, markers standing together keep the deepest cut any of them records, and a paragraph join gives its cut to markers kept from content made after the split.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A piece a kept range marker names goes with the change that took its seam; a pending paragraph mark beside the marker takes none. A pending paragraph mark waits beside kept markers across range boundaries between the paragraphs and from inside a record at the paragraph's edge. A seam's join stays on the seam when the pending insertion facing it refuses it.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An emptied container stays apart from accepted content whose insertion joined it to nothing at that edge, and rejecting an insertion that keeps nothing cancels the join of later content made inside it, so both batch orders leave the same containers.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A range marker standing at a seam without kept provenance joins the pieces around it only when it records the cut it was kept in: an accepted or authored marker keeps link and content control pieces apart in every resolution order. Kept cuts stay on markers a pending change still meets and are cleared at every depth once none does.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A kept range marker at a paragraph edge keeps its cut while the paragraph mark there is pending, so every resolution order records the same kept cuts.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record on a kept range marker the cut it stands in while a pending change keeps the pieces of a link or content control apart; the record is review state.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Kept markers held in a rejected deletion float as of the revision their cut names they were kept from.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A kept range marker's cut names the side whose record is a piece of it (`piece`), set where a pending change takes the marker's seam; accepting that change reads the markers' cut instead of the change's own kept date. Review history version 2 carries it on the marker, and version 1 files mark the markers their changes name.

  Kept markers a seam deferred on pending insertions go into no piece once those are rejected when the records the seam meets are not alike: they stand where they are, as rejecting the insertions first leaves them.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Tracked changes no longer record the date and count of kept range markers (`keptFrom`, `keptCount` on resolution joins and deferred seam groups): resolution reads each marker's own cut. Review history version 2 never writes them; version 1 files are still read, and the markers they name are marked on open.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Kept markers keep their version 1 cuts in a paragraph whose LISTNUM field folds into its marker, keep the resolution's record when a seam marks them, keep naming a piece while the change they wait on stands at the near edge of a record beside them, settle without an argument limit on long marker runs, and treat an accepted insertion whose content all goes as no piece; the wire schema checks a kept cut with the model's own validator.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A seam waits on pending insertions that join nothing themselves only when alike pieces stand beyond them, and kept markers go with the facing insertion made inside their record, so both batch orders place them alike.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Removing a pending change reads the kept markers beside it that name its side as their cut's piece, instead of its own kept date. A deferred seam's blockers name their side, or both sides, on the markers standing there, and version 1 files mark the markers a deferred seam's kept date names. A deferred seam no longer records a kept date or count, and placing kept markers once its blockers are accepted reads the markers; they stand where they are when the same resolution empties the piece across the blockers.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A removed change whose cut pieces go with it in the same resolution cuts the records left beyond them only as far as they are alike, so a range marker it keeps stands beside an unrelated container rather than inside it, as removing the pieces first leaves it.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Hand the range boundaries a rejection keeps from new content joined to nothing to a pending insertion beside them whose join cuts the record it was made in, so accepting that insertion places them as accepting it first does.

- [#2785](https://github.com/stella/folio/pull/2785) [`867366e`](https://github.com/stella/folio/commit/867366e7eedadda107ed433e73d538e01d638f43) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting the pieces of an inserted link before an insertion made inside the link, and then that insertion, leaves one link, as accepting all of them at once does. Rejecting a bookmark inserted inside the link while the link's pieces are still pending leaves the same result as rejecting it after accepting them: the pieces join around the bookmark. A resolved piece's join, or a rejection's seam, that meets a pending insertion made inside its record passes to that insertion's facing side.

- [#2953](https://github.com/stella/folio/pull/2953) [`251c0eb`](https://github.com/stella/folio/commit/251c0eb514cf2054926859f63c3a5aca09eb56db) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop deferring an accepted or kept revision's seam on pending insertions made before it: they are pieces of the record it cut, so rejecting one later keeps a range boundary it held where it was instead of moving it into the record across the seam.

- [#3253](https://github.com/stella/folio/pull/3253) [`d489af3`](https://github.com/stella/folio/commit/d489af34ca67cfe3e930177e0d87710998ee5185) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A range marker kept by a rejection never lands inside a pending tracked change; it stands beside it, and a seam waiting on pending insertions waits across such markers. Rejecting content made inside a piece of a pending insertion cancels that insertion's join facing it, and rejected pieces of one insertion meet around the markers standing between them, so the result no longer depends on resolution order. A marker kept again keeps the revision it was first kept from and the deeper cut.

- [#3156](https://github.com/stella/folio/pull/3156) [`b5bcbff`](https://github.com/stella/folio/commit/b5bcbff92b26157e7aff9ef87a67f041d3e16f70) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejoin link and content control pieces around kept range boundaries in any rejection order: count kept boundaries a pending deletion holds one by one, let a pending deletion holding only kept boundaries hold the seam between the pieces, and read past rejected new content joined to nothing when placing the boundaries.

- [#2546](https://github.com/stella/folio/pull/2546) [`56e931b`](https://github.com/stella/folio/commit/56e931bb3d7a6a9502d05422691bb1f92c00deed) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An operation finds the paragraph it addresses through an index of which top-level blocks hold each paragraph id, kept for each version of a story's block list and derived from an earlier version along the blocks that differ, instead of locating every paragraph of the story on each lookup.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving tracked insertions that split one another gives the same records in any order. A pending insertion's cut is read between the records its own pieces stand between, past insertions made after it. Two pieces of an insertion meet once accepted or rejected, whatever stands beyond them, and as deep as their own records go. Range markers kept from a later insertion beside a cut stand in it with the cut's own; one kept from outside the record a later insertion was made in stays beside that insertion. An accepted deletion of one piece leaves a cut with the other piece.

- [#3240](https://github.com/stella/folio/pull/3240) [`12980bb`](https://github.com/stella/folio/commit/12980bba330a657dca04337dd6e17c578d91490b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Look up destination IDs and range markers only when pasted content needs them.

- [#3139](https://github.com/stella/folio/pull/3139) [`a66b365`](https://github.com/stella/folio/commit/a66b365f8274e8c974a69ad673a9ffb1aff96485) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Collect newly recorded revision IDs from changed story blocks.

- [#2712](https://github.com/stella/folio/pull/2712) [`9765bb7`](https://github.com/stella/folio/commit/9765bb7dba7bf1f5291d415d3c541d63dab7524d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Setting and removing a hyperlink in the canonical session now works in paragraphs that hold content controls, direction and smart tag wrappers, equations and move markers. A content control or wrapper the range cuts stays one record and the link goes inside it; an equation in the range stays in place, unlinked, between two pieces of the link, which a later link removal treats as one. A suggested link change that would copy text inside a content control or wrapper is refused under the `linkSuggestionShape` rule.

- [#2764](https://github.com/stella/folio/pull/2764) [`5c6713e`](https://github.com/stella/folio/commit/5c6713e46711ca89329016937b20eee06c5fc553) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Undo and redo of a link edit in the canonical session show the selection its publish showed. A suggested link change over text the same author suggested inserting ends its selection at the end of the changed text, and a publish shows only a selection its addresses place exactly, moving one that falls inside a character to the nearest place an address names.

- [#2936](https://github.com/stella/folio/pull/2936) [`961d235`](https://github.com/stella/folio/commit/961d235da613ec22ff618be1f9fa3e68223a5930) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record a suggested link over a picture or shape. The suggestion strikes the drawing and inserts a linked copy, as a paste copies it; paragraphs inside a copied text box take fresh ids.

- [#2944](https://github.com/stella/folio/pull/2944) [`0f5786e`](https://github.com/stella/folio/commit/0f5786e00d784cab8ab54d66001b517712efcf1e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record a suggested link over a footnote or endnote reference. The linked copy names a copy of the note, as a paste does; resolving the suggestion removes whichever note is left without a reference.

- [#2550](https://github.com/stella/folio/pull/2550) [`d289c7b`](https://github.com/stella/folio/commit/d289c7bc435bcaad0443e712ab80fd5634a5fa06) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A list normalization replaces the paragraphs whose rendering it changed through `withParagraphRenderings`, a docx-core helper that refuses a replacement differing in anything but its list rendering and the formatting fields that rendering derives, and keeps a document known to meet the operations contract known to meet it. The next edit's operation no longer checks the whole document against the contract.

- [#3213](https://github.com/stella/folio/pull/3213) [`3f27765`](https://github.com/stella/folio/commit/3f27765a0baf4e08c00f37ce322e9e25c6e79302) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read inherited typing formatting from the caret's joined paragraph neighbors.

- [#3235](https://github.com/stella/folio/pull/3235) [`057bd5e`](https://github.com/stella/folio/commit/057bd5ecc4a3ce01a552c79c97f0e833557bc798) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Locate the addressed paragraphs of a contiguous tracked deletion and its joins without enumerating the whole story.

- [#3128](https://github.com/stella/folio/pull/3128) [`3b7ec3b`](https://github.com/stella/folio/commit/3b7ec3b8145d37878a08e801da3ef9a3f24965af) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use indexed paragraph anchors and check the final mark at the insertion boundary.

- [#3237](https://github.com/stella/folio/pull/3237) [`004b2c6`](https://github.com/stella/folio/commit/004b2c646bc05dc63b67446d3ba8bcf6a8d0f3ce) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use indexed paragraph lookup to locate the retained endpoint after a multi-paragraph deletion.

- [#3439](https://github.com/stella/folio/pull/3439) [`9cf67cb`](https://github.com/stella/folio/commit/9cf67cb8a0fc2a07fd7990445ebae37fd1dbca0d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph mark before a table resolves in a container that ends with a block content control holding its final paragraph, as it does where the container ends with that paragraph itself: a tracked or suggested paragraph inserted before the table applies, and rejecting it restores the original rather than being refused.

- [#3014](https://github.com/stella/folio/pull/3014) [`83bf0ef`](https://github.com/stella/folio/commit/83bf0ef5d732290e857270a0b51ea62011fb6f0d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Markdown export writes a tab as a tab, and a paragraph that starts with a tab or four or more blanks no longer turns into an indented code block; Markdown import reads a tab as a tab element.

- [#3195](https://github.com/stella/folio/pull/3195) [`b02302b`](https://github.com/stella/folio/commit/b02302b476dbb75e139cf92b357f3f8eaad37763) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving a row or cell the decision removes resolves the paragraph marks still pending in it first, so rejecting a suggested table whose cell ends in an inserted paragraph mark no longer refuses.

- [#2632](https://github.com/stella/folio/pull/2632) [`1352c1e`](https://github.com/stella/folio/commit/1352c1e8e584775092596133fd2ef3e81a6f5130) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting some of a run's tracked changes and rejecting others gives the same document in more orders. A pending seam no longer names an insertion that resolving the changes inside it removed, which made the next resolution refuse; and a range boundary an earlier rejection kept at a seam goes inside the piece the seam later joins.

- [#3346](https://github.com/stella/folio/pull/3346) [`1c1946e`](https://github.com/stella/folio/commit/1c1946e10a0eabe5e4bc0cf3558ead4cfb67993b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Moving a container's last paragraphs away retires the break before them on the canonical session, as deleting them does, so accepting the move leaves none of their words behind.

- [#2682](https://github.com/stella/folio/pull/2682) [`483553c`](https://github.com/stella/folio/commit/483553cce090c1bf94f0ce12e55079c62901173b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting an insertion that holds another reviewer's pending deletion of a bookmark or other range boundary now leaves that deletion pending, holding only the boundary, where the insertion stood. Accepting the deletion later removes the boundary, as deleting it directly does, and rejecting it keeps the boundary, so the result is the same whichever change is resolved first, in the engine and in the editor. A nested deletion holding no boundary still goes with the insertion.

- [#2952](https://github.com/stella/folio/pull/2952) [`8ce4fcd`](https://github.com/stella/folio/commit/8ce4fcd1350a70fdcec6b05dc1e84489218b062a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the rest of a seam on the pending deletions or move sources holding its pieces below the records it joins, with kept boundaries between them, so rejecting them later rejoins those pieces around the boundaries as rejecting everything at once does.

- [#3156](https://github.com/stella/folio/pull/3156) [`b5bcbff`](https://github.com/stella/folio/commit/b5bcbff92b26157e7aff9ef87a67f041d3e16f70) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a rejected change's seam inside a content control or link on a pending insertion standing between the pieces, so rejecting that insertion later rejoins them as rejecting both at once does.

- [#3257](https://github.com/stella/folio/pull/3257) [`d227b09`](https://github.com/stella/folio/commit/d227b09b7c70f955ef415e404a13da05798edd51) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving the removal of a row holding a nested table settles each of its paragraph marks once, so accepting a document deleted over a nested table no longer refuses.

- [#3325](https://github.com/stella/folio/pull/3325) [`2bad3f5`](https://github.com/stella/folio/commit/2bad3f548dfc81d52d2795e7f3c09b6f6b1e22a1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a note reference's deletion also rejects the note deletions its author made with it.

- [#2800](https://github.com/stella/folio/pull/2800) [`ef7a2dc`](https://github.com/stella/folio/commit/ef7a2dc7ac71f8307c9279321b5ef7abc3312e36) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A new note takes its place in its note list by id, before the first note of a greater id, rather than at the end. Notes added concurrently by several peers now leave the same note list whichever is added first.

- [#2794](https://github.com/stella/folio/pull/2794) [`c12559e`](https://github.com/stella/folio/commit/c12559e5583c13ac49ee91b1d06a6a630bf81a8b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving a tracked change removes a footnote or endnote whose only reference sat inside a field's result that the change removed, and keeps a note that a field's result elsewhere still references. Every reading of where a document references a note, after any edit as well as after a resolution, enters field results at any depth.

- [#3347](https://github.com/stella/folio/pull/3347) [`1b516f0`](https://github.com/stella/folio/commit/1b516f0f7e70e709dfce538c17c1e943813fbaef) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A canonical tracked operation deleting a note's last live reference also deletes the note's text and paragraph marks, as operations of their own; a comparison states each note's changes as a story of its own. docx-core adds `noteDeletionOps`.

- [#2806](https://github.com/stella/folio/pull/2806) [`b9566e6`](https://github.com/stella/folio/commit/b9566e6a660eb3aa6a843611237a5dcbfdad5dca) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Creating a numbering instance puts the instance, and the abstract numbering it brings, before the first entry of a greater id instead of at the end, so concurrent creations leave one order whichever applies first. A creation that restores a captured numbering part, such as undoing a deletion, puts them back where they were: undoing the deletion of an instance that was not the last one is no longer refused.

- [#3489](https://github.com/stella/folio/pull/3489) [`881fa93`](https://github.com/stella/folio/commit/881fa93e43ca47a736fcab0ec069d7d7a3d784d1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A document holds one range per move-range id. Inserted content leaves out a copy of a move range boundary that stands in the document, in every mode and on paste, as a copy of moved text is no part of the move; content pasted over its own source, or put back by an undo, keeps its boundaries. A part that repeats a range is repaired on open: its first range is kept, and each further one is dropped and reported as `dropped-move-range`.

- [#3156](https://github.com/stella/folio/pull/3156) [`b5bcbff`](https://github.com/stella/folio/commit/b5bcbff92b26157e7aff9ef87a67f041d3e16f70) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Put the range boundaries a rejected insertion keeps inside the link or content control piece its one-sided join continues, as rejecting it together with the change that left them does.

- [#2773](https://github.com/stella/folio/pull/2773) [`8517530`](https://github.com/stella/folio/commit/85175307e3323d52cf6ee9cc5f557442193ed733) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A shape outline written from the model, such as one read from a VML text box, is written as an empty element when it has no child elements, so saving an unedited document a second time gives the same bytes as the first. A line end whose `type` the source omits keeps no type in the model and is saved without one.

- [#3175](https://github.com/stella/folio/pull/3175) [`d489abb`](https://github.com/stella/folio/commit/d489abb0eef6dfbe9370d6f6fa5e6455029f72d5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of one's own inserted paragraph break right before a table whose rows are all pending insertion records the deletion beside that insertion and takes the paragraph's words, so accepting it removes the paragraph as the direct deletion does.

- [#2535](https://github.com/stella/folio/pull/2535) [`063e63a`](https://github.com/stella/folio/commit/063e63a6a5a26da7e62220bc7884518032eecd16) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `withParagraphCopies` puts copies of a story's paragraphs, at any depth, in their places, its section views included. Each copy must hold the same value under every field as the paragraph it replaces, so a document known to meet the seed contract stays known to meet it and the next operation skips the full check.

- [#2792](https://github.com/stella/folio/pull/2792) [`1439897`](https://github.com/stella/folio/commit/143989737ed1513703f19b6a0c61ad1a0d5c3d0d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Paragraph properties an edit sets read in the live model as the saved file reopens: the spacing sides the paragraph states are marked as the parser marks them, and a hanging indent is held as the parser reads it, a negative first line flagged hanging. `paragraphSpacingExplicit` is the one rule for the spacing sides, used by the parser and by the patch an operation states; paragraphs no operation names keep their formatting as it is.

  A tracked paragraph property change records the properties the paragraph now has as its current formatting, as a reopened file reads them from the paragraph itself; folding a later change keeps the first change's previous properties.

- [#3103](https://github.com/stella/folio/pull/3103) [`9a03f56`](https://github.com/stella/folio/commit/9a03f5659cfe03d258c8be9795306e162acdc3df) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The package paragraph-id census reads each record once, so a batch of operations no longer walks every paragraph of the package per operation.

- [#2634](https://github.com/stella/folio/pull/2634) [`f6ccd85`](https://github.com/stella/folio/commit/f6ccd85f6d3d3546b685ab43482c35170d633d4e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a paragraph mark and the insertions beside it at once now gives what rejecting them one at a time gives: a range boundary kept at the seam the paragraphs join at goes inside the piece the join merges, instead of keeping the pieces apart.

- [#3067](https://github.com/stella/folio/pull/3067) [`55d8f81`](https://github.com/stella/folio/commit/55d8f81e5b7dd1b5c6ab13aa9cefa0a48f0ec384) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph mark can hold another author's pending insertion and a later pending deletion of it. Both records are read, written in schema order, listed as changes and resolved one at a time or together; deleting another author's inserted paragraph records the deletion beside the insertion instead of removing the paragraph.

- [#3021](https://github.com/stella/folio/pull/3021) [`06170f1`](https://github.com/stella/folio/commit/06170f1f7ada907cf163c52e62178e97ab522b69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve a tracked column insertion or deletion that only some rows take: a grid column goes only when every row loses its cell, and a row that loses a cell in a kept column widens the cell before it (or the one after it, at the start of the row), instead of refusing the review or the save.

- [#3457](https://github.com/stella/folio/pull/3457) [`fd15f56`](https://github.com/stella/folio/commit/fd15f565213b20a2cd3bc72d16b63750ecc0dc83) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Deleting a paragraph break after an empty paragraph that a pending join runs words into keeps that paragraph's properties, in direct, tracked and suggested edits alike, so accepting a tracked deletion of a last paragraph leaves what the direct deletion does. A tracked deletion that retires the break before a container's pending end records those properties too, and one after a break already pending deletion leaves the last paragraph's properties as they are.

- [#2825](https://github.com/stella/folio/pull/2825) [`06d77b2`](https://github.com/stella/folio/commit/06d77b2143fcd9a5c3c257475686c2bdb36fac5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked paragraph property change records, as the paragraph's current properties, what the paragraph states in its own properties, as a reopened file reads them. A direct property edit of a paragraph with such a change pending keeps the change recording them, and undoing that edit restores the change exactly.

  List normalization restores derived formatting and pending paragraph changes atomically during undo.

- [#2850](https://github.com/stella/folio/pull/2850) [`bf33260`](https://github.com/stella/folio/commit/bf33260ccc34df59ce99454a96ebedb15c462268) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retire point-comment threads when their last anchor is removed while preserving source orphans and exact undo.

- [#3217](https://github.com/stella/folio/pull/3217) [`4551ebf`](https://github.com/stella/folio/commit/4551ebf8deadb86764958bbd98f7ef5864a039d5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked paragraph-property edit that returns a paragraph to the properties its pending change started from now removes that change instead of keeping one that records nothing.

- [#3209](https://github.com/stella/folio/pull/3209) [`cb92ea8`](https://github.com/stella/folio/commit/cb92ea8a895f4e101e0a1c05cb01c11e2a1232d8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prepare supported tracked paragraph insertions through native validation before materializing the story.

- [#3234](https://github.com/stella/folio/pull/3234) [`3bc218e`](https://github.com/stella/folio/commit/3bc218eddc1f91dcad2715304b29bae66e09fbfb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Avoid gathering package identities when a native range deletion needs no extra IDs.

- [#3143](https://github.com/stella/folio/pull/3143) [`7f8a76b`](https://github.com/stella/folio/commit/7f8a76bea5180f3ca3e37650c6f20bd2c6c65708) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse indexed paragraph lookups and skip identity collision scans when tracked change resolution leaves no incoming records.

- [#3190](https://github.com/stella/folio/pull/3190) [`cf17a09`](https://github.com/stella/folio/commit/cf17a09a2ec8a9029f89e1ec26db8e01bcce7121) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Skip scalar descendants and intermediate arrays during document identity censuses.

- [#2886](https://github.com/stella/folio/pull/2886) [`0de6793`](https://github.com/stella/folio/commit/0de6793d56995d2c1ba7f91ec1ae9fb0491b5967) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve opaque inline XML when resolving paragraph marks, including captures with no visible text.

- [#2986](https://github.com/stella/folio/pull/2986) [`f0c1fc5`](https://github.com/stella/folio/commit/f0c1fc53a71bcf06847241cf2f2f2d1be3e65f49) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse tracked edit batches that fail to restore the original on rejection or lose protected content on acceptance.

- [#2866](https://github.com/stella/folio/pull/2866) [`a33e505`](https://github.com/stella/folio/commit/a33e505587e7e7210a748649cdac6e0d40c9b2b6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse concurrent erasures that can reassign tracked deletion fragment identities, and report dependent pending review decisions when their creator is dropped.

- [#3187](https://github.com/stella/folio/pull/3187) [`61db1c8`](https://github.com/stella/folio/commit/61db1c84cfce1173e7de83ad6c0bdee85c956268) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A table whose rows end early is read with the empty columns they imply after them, so its revisions resolve; a column only such a row's `gridBefore` or `gridAfter` holds stays when the cell in it is removed, and the cell beside it takes the removed cell's column. Column edits still need every row to cover the grid.

- [#3449](https://github.com/stella/folio/pull/3449) [`25386cf`](https://github.com/stella/folio/commit/25386cfe28589d2ef64439bbc6d21bb34bd60ba5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read range boundaries standing between blocks as positions, not blocks: deleting a paragraph finds the paragraph or table past them, or ends its container, and a paragraph that joins the next one takes them in where the two meet. Resolving a paragraph mark before a table or block content control past them removes or keeps the paragraph as it would without them.

- [#3496](https://github.com/stella/folio/pull/3496) [`4d4fc90`](https://github.com/stella/folio/commit/4d4fc907545fcf58534ac2382db9bfb69e8bf041) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Plan text ranges across range boundaries standing between paragraphs as one run of paragraphs: replacing or deleting such a range joins them with the boundaries where they meet, directly or once its tracked deletion is accepted, and a direct deletion's caret lands in the paragraph past them.

- [#2587](https://github.com/stella/folio/pull/2587) [`1212c0b`](https://github.com/stella/folio/commit/1212c0bcb7b03c585300b72fead8cdb95518048c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The style resolver and the style engine take style definitions typed read-only and hand out read-only results: `getStyle`, `getParagraphStyles`, `getTableStyles`, `getDocDefaults`, the default style getters, `resolveParagraphStyle`, `resolveParagraphStyleInTable`, `resolveRunStyle` and `getRunStyleOwnProperties`. `getRunStyleOwnProperties` returns the style's own run properties instead of a shallow copy. The text and paragraph formatting merges keep plain inputs plain and give read-only results for read-only inputs, and a merge with a source is a new record. The paragraph attributes the style cascade fills (`borders`, `shading`, `tabs`, `defaultTextFormatting`, `_tableRunFormatting`, `_resolvedFormatting`, `_styleResolvedFormatting`), a table cell's `borders` and `_resolvedBorders`, and the selection snapshot's `textFormatting` and `paragraphFormatting` are read-only. `writableCopy` in `@stll/folio-core/types/documentModel` is the copy a writer takes before it changes such a value. `DeepReadonly` moves to `@stll/docx-core/model`, and core re-exports it. React's `PagedEditor` reads the model's styles without `readModel` and passes them on read-only; `getParagraphTabs` and React's and Vue's ruler tab stops are read-only. Saving a paragraph's or a table cell's borders, shading and tab stops copies them, so the saved document owns what it states.

- [#3227](https://github.com/stella/folio/pull/3227) [`f8a77b6`](https://github.com/stella/folio/commit/f8a77b64fc6c9490de15825aafb58e8ba9301f24) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use the matching quote to locate XML tag boundaries in long attribute values.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A removal's seam waits on the insertions standing between pieces of a pending insertion only, not of a pending deletion, and a removal resolved with the pieces it cut reads that cut no deeper than removing the pieces first leaves: a range marker it keeps lands in the same place whichever order the neighbouring changes are resolved in.

- [#2753](https://github.com/stella/folio/pull/2753) [`9a525cd`](https://github.com/stella/folio/commit/9a525cdfe08bb09b675c19393ef09581da99af1b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refresh the canonical WebAssembly artifacts for the updated compatible Rust dependencies.

- [#2533](https://github.com/stella/folio/pull/2533) [`0e5d72a`](https://github.com/stella/folio/commit/0e5d72a30efb7200990b07fd6facfa34c1d7c819) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Applying an operation refuses an edit whose document would hold one paragraph id twice, with the seed contract's `duplicateBlockId` reason. The check reads a paragraph id index of the package, derived where it can from a package indexed earlier, so it usually walks only what the edit changed.

- [#3146](https://github.com/stella/folio/pull/3146) [`d29dc0c`](https://github.com/stella/folio/commit/d29dc0ccd9b044f6dd33022fd3e9e0231239ef28) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The legacy engine no longer opens a package whose XML names an undeclared entity (`&nbsp;`, `&copy;` or any name but XML's five built-ins); it used to read common HTML names as their characters and keep other names as written. Both engines now refuse such a package with `UndeclaredXmlEntityError`, which names the part, the reference's offset and the entity. Character references are read as before. The CLI's save check refuses such a part as not well-formed.

- [#3457](https://github.com/stella/folio/pull/3457) [`fd15f56`](https://github.com/stella/folio/commit/fd15f565213b20a2cd3bc72d16b63750ecc0dc83) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a paragraph inserted before a block content control or custom XML that ends the body, holding its last paragraph, or before another block in a body whose last paragraph a bookmark follows, removes the paragraph whole instead of leaving an empty one, so a document saved with a pending suggestion there is the original.

- [#3359](https://github.com/stella/folio/pull/3359) [`dad63bd`](https://github.com/stella/folio/commit/dad63bd79f2badcb1cc714092e45e5c050244787) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a section-property change whose previous record is empty removes the section, rather than leaving a section with default page settings.

- [#2553](https://github.com/stella/folio/pull/2553) [`076e70d`](https://github.com/stella/folio/commit/076e70d424946ba0c00259d6c2aa2ea9bb112f1c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving revisions places the range boundaries that removed content leaves by the records that remain: rejecting content that cut a pending insertion joins the insertion's pieces again, and a wrapper resolved in the same pass no longer counts as a level, so resolving revisions at once or one operation at a time gives the same document.

- [#3359](https://github.com/stella/folio/pull/3359) [`dad63bd`](https://github.com/stella/folio/commit/dad63bd79f2badcb1cc714092e45e5c050244787) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting an inserted paragraph mark that ends a section takes that section with it instead of moving it to a neighbouring paragraph.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting an insertion that keeps nothing leaves content made inside it with the join that insertion had to the record beyond it, as rejecting it after accepting that content does. It still cancels the join the content had toward the rejected insertion itself.

- [#3499](https://github.com/stella/folio/pull/3499) [`11e8e93`](https://github.com/stella/folio/commit/11e8e93223c602576488ce82264348051cf4fa6f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove an added paragraph before a content control or preserved block in whichever order its break and words are rejected.

- [#3393](https://github.com/stella/folio/pull/3393) [`3fabdf7`](https://github.com/stella/folio/commit/3fabdf7e7948be29df642f4a29a2530f7d468aca) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a paragraph-property change keeps every indentation the restored properties state, in both sessions, including a value equal to the one the paragraph's numbering supplied before. `registerStatedParagraphFormatting` and `paragraphStatesItsFormatting` mark and read formatting an edit gives as stated; `transferParagraphPropertySnapshot` keeps a paragraph's registration on a copy of it, so a property change recorded on the copy records the properties it states.

- [#3467](https://github.com/stella/folio/pull/3467) [`34d621a`](https://github.com/stella/folio/commit/34d621a90c964a1d8491f221510498b61fdd34e4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a paragraph split together with the paragraph's own inserted break and a pending property change, before a block content control or another block, restores the paragraph's properties: a paragraph joining it in the same resolution does not stand for one that stays, also with a range boundary between them.

- [#2854](https://github.com/stella/folio/pull/2854) [`af6aa54`](https://github.com/stella/folio/commit/af6aa54464e2e2775e9e9a90830b8cc3c86edfa5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejoin marker-only source record pieces when rejecting a tracked insertion across their recorded cut.

- [#2593](https://github.com/stella/folio/pull/2593) [`bf37d28`](https://github.com/stella/folio/commit/bf37d280c8fee4b83feca7e529722f64afb5fe4f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - When rejecting a revision joins the pieces of another pending insertion again, the seams those pieces were waiting on name the joined record, so rejecting the rest in any order resolves cleanly.

- [#2625](https://github.com/stella/folio/pull/2625) [`1a1a311`](https://github.com/stella/folio/commit/1a1a311e47af052a77027780826f164ac2157fdd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting an insertion together with the pieces of a change it split no longer merges the content controls or links beside them, matching the result of rejecting the changes one at a time.

- [#2821](https://github.com/stella/folio/pull/2821) [`68b185a`](https://github.com/stella/folio/commit/68b185a20d5549e4255eeee6580f1cc65b5fc954) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting a deletion inside a link and then an insertion of new content beside it, with the rest of the link after that insertion, leaves the same records as accepting both at once: the inserted content stays its own record. Rejecting the insertion still joins the link's pieces around the deleted text.

- [#2633](https://github.com/stella/folio/pull/2633) [`991f517`](https://github.com/stella/folio/commit/991f517792615d81303ecf155aa1a37ce1550dd8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting an insertion together with later insertions beside it joins the text it split again, as rejecting the changes one at a time does.

- [#2803](https://github.com/stella/folio/pull/2803) [`1326d7c`](https://github.com/stella/folio/commit/1326d7c68c20bebd8763d2f1c8f493244fe712a8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting a tracked deletion of a whole paragraph no longer leaves its bookmarks or move ranges standing between blocks. Every bookmark, move range boundary and comment anchor the paragraph held goes to the nearest paragraph that stays, as deleting the paragraph directly leaves them, and the edit reports that paragraph as changed. Before a table with no paragraph ahead of it, they run into the table's first cell, as comment anchors already did, and where the table cannot take them the paragraph stays holding them. The editor resolves paragraphs the same way. `paragraphRemains` is no longer exported from `@stll/docx-core/ops`.

- [#3383](https://github.com/stella/folio/pull/3383) [`6929864`](https://github.com/stella/folio/commit/692986472ca8c00e0b98db80727e8d224fcc69a4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open packages with dangling or unpaired comment range markers, or comment paragraphs repeating a paragraph id, on both engines by repairing them on open and reporting each repair as a parse warning.

- [#3010](https://github.com/stella/folio/pull/3010) [`5956025`](https://github.com/stella/folio/commit/59560253e48152663d16f6622a7f2b6d73a15aa6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore missing DrawingML namespace declarations in imported drawings.

- [#3383](https://github.com/stella/folio/pull/3383) [`6929864`](https://github.com/stella/folio/commit/692986472ca8c00e0b98db80727e8d224fcc69a4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Opening a package repairs a comment range boundary with no partner between a table's rows or a row's cells, where a reference anchors the comment, and a comment's repeated references (new parse warning `duplicate-comment-reference`); both sessions open such packages.

  Comment range boundaries pair only where the pairing is unambiguous, so a lone start never widens a full range; a comment keeps its first full range and reports a later one with its text (new parse warning `dropped-comment-range`); a comment whose only anchor sits between table rows or cells becomes a point at the nearest paragraph.

- [#2554](https://github.com/stella/folio/pull/2554) [`a103613`](https://github.com/stella/folio/commit/a1036134376f62cb6a1d45195e6b7d9e2d5b9fee) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Text replacing a range inside an inline wrapper or content control continues it after every leaf the deletion keeps, comment anchors included, so a direct edit matches the accepted tracked one.

- [#3456](https://github.com/stella/folio/pull/3456) [`062f0c9`](https://github.com/stella/folio/commit/062f0c9ec37ad073b7b70c6087f1440936f8ef69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Text replacing every word a comment covers goes past the comment's end, and past other ranges' ends standing before it, while staying in the link, content control or wrapper the words were in. A direct replacement takes the comment, and a link left holding only its markers goes with them; a tracked one keeps the comment on the struck words, and accepting it gives the direct edit.

- [#3264](https://github.com/stella/folio/pull/3264) [`7d5ee02`](https://github.com/stella/folio/commit/7d5ee02e90d536fc2c6e8acee2c3da050d3ab0a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A replacement whose formatting patch clears everything writes a plain run, which joins the identical run beside it instead of standing apart with an empty property set.

- [#3058](https://github.com/stella/folio/pull/3058) [`34a05ca`](https://github.com/stella/folio/commit/34a05caddaaa33013d8a6c73907806a588904fcc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Put the words that replace a range past the end of any comment the range reaches, so a replacement no longer stretches a comment over new words: tracked, the struck words keep their comment ranges; direct, a comment whose words all went goes with them. A direct replacement across a comment's end no longer refuses.

- [#2921](https://github.com/stella/folio/pull/2921) [`657e51d`](https://github.com/stella/folio/commit/657e51d03686ae4ff65e9d49bf7a1eb079390ddc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Place a reply's comment markers beside its thread's markers inside the hyperlink, content control or wrapper they stand in. A reply no longer cuts a link into pieces, no longer needs new ids to split a content control, and no longer leaves a wrapper the canonical editor cannot project.

- [#3438](https://github.com/stella/folio/pull/3438) [`7717816`](https://github.com/stella/folio/commit/77178167741d0c5fff0347cf17267cefb5fd859c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An accepted insertion whose content the same resolution removes cuts nothing between its neighbours, as resolving that content first leaves them. Accepting the deletion of an inserted paragraph break rejoins the records its split cut, as rejecting the insertion does. A paragraph with no content of its own, removed before a table or block while the paragraph before it joins in the same resolution, leaves all it held in that paragraph, as resolving the marks one at a time from the first does. A kept range marker whose named piece of its cut an accepted deletion takes goes into the piece across from it, as rejecting the revision it was kept from after that deletion places it.

- [#2849](https://github.com/stella/folio/pull/2849) [`4372778`](https://github.com/stella/folio/commit/4372778ef35a51ec4423d6073328bc13f7844630) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retain source paragraph marks when tracked insertion before a table cannot resolve an inserted mark, so suggestion rejection restores the original container.

- [#3138](https://github.com/stella/folio/pull/3138) [`c043eee`](https://github.com/stella/folio/commit/c043eee5a1e81acf0ea2e0ebb882f858f5e3eea3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve paragraph absorption chains in linear time while preserving survivor IDs and receipt order.

- [#3183](https://github.com/stella/folio/pull/3183) [`78f418a`](https://github.com/stella/folio/commit/78f418a49d0fa768528e559cd29263b5e831aedd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Revision resolution reaches the tracked property changes of a body section that ends in a paragraph. `resolveSectionChange`, `resolveSectionChanges`, `sectionRejectProperties` and `sectionResolutionPatch` move to `@stll/docx-core/ops`, which the reviewer now calls.

- [#3328](https://github.com/stella/folio/pull/3328) [`27fa79e`](https://github.com/stella/folio/commit/27fa79e8e59c55ace543334ff176708439d5a1a9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve a decision that takes every cell of a table away (deleted cells accepted, inserted cells rejected) by removing the table, as removing its last row does, instead of refusing it.

- [#3430](https://github.com/stella/folio/pull/3430) [`40b24de`](https://github.com/stella/folio/commit/40b24de106a4568b4215ce311a67807add5018e0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving a paragraph mark with bookmark or other range boundaries standing after it now resolves as it does with nothing between. Before a block content control or custom XML, an empty paragraph goes and the boundaries stay, and still-pending words run into the control's first paragraph with the boundaries where the two meet. Before only boundaries at the container's end, the paragraph stays, so the container still ends with a paragraph.

- [#3275](https://github.com/stella/folio/pull/3275) [`60fa1a4`](https://github.com/stella/folio/commit/60fa1a42ae25c00015618897e6c2873851626c8d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Revision resolution reaches a tracked section-property change carried by a paragraph of a header, footer or note: it resolves on that paragraph by the section calculation body sections use, instead of being refused as unreachable.

- [#2786](https://github.com/stella/folio/pull/2786) [`874b995`](https://github.com/stella/folio/commit/874b99527e3642d07b0177488f192314952fd98d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `resolvedDocumentView` resolves a story that ends in a table with the closing paragraph a session reads it with, and resolves a terminal-table review carrier as its table does; `storyLacksClosingParagraph` states that ending for both.

- [#2810](https://github.com/stella/folio/pull/2810) [`7c475ee`](https://github.com/stella/folio/commit/7c475eea28a89167cf8d3eba9916d1f7529baa69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph restoration, such as undoing an edit, and a concurrent new comment anchored in that paragraph now meet by the restoration's span: when neither end of the comment's anchor lies inside the stretch the restoration changes, both are kept in either order instead of one being refused because the other changed the paragraph.

- [#2752](https://github.com/stella/folio/pull/2752) [`99824b8`](https://github.com/stella/folio/commit/99824b8765c0ee279fd6944a26ffd37ea77e69df) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph restoration (an undo that gives a paragraph back its content, and each paragraph a comment restoration rewrites) no longer conflicts with every concurrent edit of that paragraph. It conflicts only with an edit reaching into the stretch it changes; any other typing, deletion, format or inline record split or join in that paragraph passes it in either order, in the body, a table cell, a header or footer, or a note.

- [#2869](https://github.com/stella/folio/pull/2869) [`168353b`](https://github.com/stella/folio/commit/168353b31c3eae6d8f605ff01e3149a436ddfada) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore original paragraph properties when rejecting a tracked deletion that cancels an earlier inserted paragraph mark.

- [#2862](https://github.com/stella/folio/pull/2862) [`54a08af`](https://github.com/stella/folio/commit/54a08afe458f3a32655961232826410dd8394da5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep concurrent edits outside restored revision boundaries without allocating new identities.

- [#2872](https://github.com/stella/folio/pull/2872) [`a687a7c`](https://github.com/stella/folio/commit/a687a7c3608cab3440798ff1d6c7060b20a0cd53) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retire exhausted comment threads when note bodies disappear and preserve their exact relationship state through undo and review resolution.

- [#3305](https://github.com/stella/folio/pull/3305) [`4d1e3ec`](https://github.com/stella/folio/commit/4d1e3ec9895e5b9836765b851f528f3fbea8ae0a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked or suggested deletion of a paragraph one's own pending insertion added takes the insertion back: before a following paragraph it is removed outright, and after the container's last paragraph the insertion is rejected, so the paragraph before keeps its identity and properties. A comment the paragraph alone carries retires with it and is reported in `retiredCommentIds`. The legacy session no longer leaves an empty paragraph-property change on that paragraph.

- [#3160](https://github.com/stella/folio/pull/3160) [`efe7932`](https://github.com/stella/folio/commit/efe79325abcd4a5ca66ef9fb8f17d7a56bed9c01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse cut-depth summaries when resolving tracked paragraph splits along the same block list.

- [#3249](https://github.com/stella/folio/pull/3249) [`f8bd99d`](https://github.com/stella/folio/commit/f8bd99d5a8cdf3767b12b5599333b1ac167a26ea) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse immutable change bounds when planning and replay replace the same paragraph slot.

- [#3232](https://github.com/stella/folio/pull/3232) [`2d91243`](https://github.com/stella/folio/commit/2d91243c54a6dac2c98cda0df83adda5342cb49a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse a successful direct preflight when planning an identical single-segment tracked deletion.

- [#3140](https://github.com/stella/folio/pull/3140) [`2474150`](https://github.com/stella/folio/commit/24741500b5f1865ed2c06d5224c95e3a6cf352c7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Update the checked trailing section directly for ordinary paragraph splices.

- [#3248](https://github.com/stella/folio/pull/3248) [`c9be747`](https://github.com/stella/folio/commit/c9be7472fbd6f034300431858798b3e6dfc5ac54) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse exact comment anchor comparisons for unchanged story content.

- [#3343](https://github.com/stella/folio/pull/3343) [`2214fbe`](https://github.com/stella/folio/commit/2214fbec80d6a9d0fffad8d19b60dd6385725a6a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a retained range marker at its seam while an earlier insertion still owns the facing review join.

- [#2846](https://github.com/stella/folio/pull/2846) [`fc62152`](https://github.com/stella/folio/commit/fc6215232837908e44bbdecd24e02d40d9efb2dc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep pending revision decisions when concurrent anchored comments do not cut revision records, preserving independent edits during client rebasing.

- [#2766](https://github.com/stella/folio/pull/2766) [`a947ffd`](https://github.com/stella/folio/commit/a947ffd65450dafb606d09195180fdf3aff4413d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Writing revision info onto a table row's tracked insertion or deletion, or onto a cell's or a table's tracked property change, now reports the paragraphs of that row, cell or table as modified, so an edit's touched blocks cover every record it changed.

- [#2685](https://github.com/stella/folio/pull/2685) [`1f655c2`](https://github.com/stella/folio/commit/1f655c22e4e7a1eae9719110924c51f73936935c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Deleting a table row directly that starts a vertical merge the next row continues hands the merge to that row, with the removed cell's properties and content, instead of leaving a continuation without a start and dropping the merged cell's words. The merge ends there when no later row continues it.

- [#2760](https://github.com/stella/folio/pull/2760) [`526eb9f`](https://github.com/stella/folio/commit/526eb9f5e8bd6be659013a60c781aa21da22a417) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Undoing a row deletion or insertion restores a cell's final paragraph mark that no longer carries its row's revision, such as after the row's insertion was accepted or re-attributed on its own.

- [#3377](https://github.com/stella/folio/pull/3377) [`76ee4cd`](https://github.com/stella/folio/commit/76ee4cdcf5eaadb5822e38e459b866e3f2f063ea) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Drop a run's pending property change when a tracked edit returns the run to the properties the change started from.

- [#3375](https://github.com/stella/folio/pull/3375) [`db4feb6`](https://github.com/stella/folio/commit/db4feb6fc425738ed645977914afc843ad103a7f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a run's pending property change when a suggestion formats the run: a suggestion folds only into its own change.

- [#2737](https://github.com/stella/folio/pull/2737) [`72341ee`](https://github.com/stella/folio/commit/72341eef4aeb81fd029340183bc988f748f03beb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Run property patches reach the runs of a field's displayed result and the run holding a note's own reference mark, so formatting a selection that holds a field or a note mark formats them too. A tracked patch records its change on each of those runs with an id of its own, and accepting, rejecting and listing revisions reach the changes there.

- [#2985](https://github.com/stella/folio/pull/2985) [`cc89cad`](https://github.com/stella/folio/commit/cc89cad71c1e8edae0efbc03b4ee14fc09a3228c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Raise the minimum supported Rust version of the `stella-docx-kernel` and `stella-text-shaper` crates to 1.100.

- [#2954](https://github.com/stella/folio/pull/2954) [`32f1feb`](https://github.com/stella/folio/commit/32f1febddabd8a1f3eac2ceed8618ee38dcecaab) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rebuild the canonical WebAssembly artifacts with Rust 1.100.0.

- [#2859](https://github.com/stella/folio/pull/2859) [`5831282`](https://github.com/stella/folio/commit/583128222fef935acfa2dd2cdfa9d25bb7b3acba) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align tracked-edit conformance classification with scoped source restorations.

- [#3462](https://github.com/stella/folio/pull/3462) [`367d51a`](https://github.com/stella/folio/commit/367d51a864a09df4fdb96c3377686a9e33f7bda5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting content made inside a pending insertion while another pending insertion stands between its pieces keeps the seam on that insertion: rejecting it later joins the pieces, or keeps its range markers in their cut, as rejecting both at once does.

- [#3500](https://github.com/stella/folio/pull/3500) [`ad43237`](https://github.com/stella/folio/commit/ad432376058969e535177a9705af5ff1999de54a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving tracked changes at a cut record leaves the same document in more orders. Insertions made one at another's edge at the end of a cut piece are part of that piece, as deep as they are joined to it, and the one nearest the seam hands its join on when it goes. A cut goes on through the pieces its own change removes, and the range markers kept there go into the records it leaves. Markers kept from a record of a rejected insertion go into the content made inside that record, and an accepted deletion that takes a split's piece takes the split's cut with it.

- [#3107](https://github.com/stella/folio/pull/3107) [`ecfa20c`](https://github.com/stella/folio/commit/ecfa20c215ae431c5ef8716877394d204c04c14d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Delete a paragraph that ends a section, tracked or directly. The section break is the paragraph mark's: a direct deletion, or accepting the tracked one, joins the sections on either side under the properties of the section after it; rejecting restores every section. One's own added mark that ends a section records the deletion beside the insertion instead of taking the insertion back.

- [#3385](https://github.com/stella/folio/pull/3385) [`eaf8da4`](https://github.com/stella/folio/commit/eaf8da49559da56e92cfbd9f51ae2b9c35032e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph whose mark ends a section can be deleted right before a table, directly or tracked. Its section break goes with it, removed explicitly with an exact inverse, so its section and the next one become one under the later section's properties; accepting the tracked deletion gives the direct result and rejecting it restores the paragraph and its section.

- [#3405](https://github.com/stella/folio/pull/3405) [`bd2a344`](https://github.com/stella/folio/commit/bd2a344c3aa8c20ee5e0b47a7ac8ef6521205b9f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A block replacement that removes a section break together with a note's only reference returns an inverse that restores the note and then the blocks with their section, instead of failing.

- [#3465](https://github.com/stella/folio/pull/3465) [`d1309d5`](https://github.com/stella/folio/commit/d1309d50ddf8ca69be0503e83e2d78fbd47e14b3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Document operations keep section breaks out of table cells. An operation after which a section ends in a cell where none ended before is refused with `sectionBoundary`, whether it makes, moves, copies or restores the break, and so is setting a section endpoint on a cell's paragraph. A break a document already held in a cell stays through edits that leave it where it is. A paste or a drop into a cell, and a pasted table, come without the section breaks of their paragraphs.

- [#2547](https://github.com/stella/folio/pull/2547) [`951b7e1`](https://github.com/stella/folio/commit/951b7e181fcab8c4572be600118e3d624e2fafb9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The canonical session judges an edit's result by the blocks the edit replaced: the blocks it kept were judged before, and the comment check runs again only when the comments, the stories, or the order, kind and id of the replaced blocks' comment anchors change. docx-core exports `commentAnchorSequence`, the anchors a run of blocks adds as the comment check reads them.

- [#3408](https://github.com/stella/folio/pull/3408) [`4bae709`](https://github.com/stella/folio/commit/4bae7096d08e0d6e4582211e28518d3e5f41eb1e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Sequencing refuses three more concurrent pairs whose orders give a record different ids. A paragraph restoration conflicts with an edit at an edge of its span when it moves where a revision or content control starts or ends there. A direct deletion conflicts with a concurrent direct insertion that supplies new ids when the deletion removes content up to or across the insertion's point. Two insertions that put in the same identified record conflict.

- [#3223](https://github.com/stella/folio/pull/3223) [`7efec2c`](https://github.com/stella/folio/commit/7efec2cc40b080cea9093df94722ba4a60daa01b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse the XML 1.0 character check when reading and validating document parts.

- [#2670](https://github.com/stella/folio/pull/2670) [`528c72e`](https://github.com/stella/folio/commit/528c72ea5a6c75f30f9d423b305132f476ba738c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove rejected empty inserted paragraphs before opaque blocks when final-paragraph and section ownership remain valid.

- [#2670](https://github.com/stella/folio/pull/2670) [`528c72e`](https://github.com/stella/folio/commit/528c72ea5a6c75f30f9d423b305132f476ba738c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore imported terminal paragraph marks through exact snapshots when undoing review resolution.

- [#2919](https://github.com/stella/folio/pull/2919) [`b70694a`](https://github.com/stella/folio/commit/b70694a24b37f0623cf73935ace88773b4a2825f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A paragraph split at the point of a concurrent suggested insertion is accepted in the authority's order instead of being refused: the edit sequenced later goes after the other.

- [#3320](https://github.com/stella/folio/pull/3320) [`e5557c0`](https://github.com/stella/folio/commit/e5557c07c121c01cc2a38f8517460ce99c688007) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A split that keeps the paragraph's identity on its first half gives that half the surviving mark's properties and leaves the paragraph's own on the second, so a split of a paragraph whose mark is pending deletion reads, once accepted, as the merged paragraph split directly.

- [#3061](https://github.com/stella/folio/pull/3061) [`9771dcf`](https://github.com/stella/folio/commit/9771dcfad77176f2b06876d933def6053b6da70e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a tracked change together with a paragraph break inserted inside it joins only the change's own seam, as rejecting them one after another in any order does: the pieces either side of the break no longer join the records beyond them at the break.

- [#2909](https://github.com/stella/folio/pull/2909) [`12df99f`](https://github.com/stella/folio/commit/12df99f441581a4601370926ea2b449ec3af9660) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejoin split paragraphs only as deep as the split still cuts once a resolution unwraps the tracked change it cut: two adjacent links inside an accepted insertion or a rejected deletion or move source no longer merge into one.

- [#2871](https://github.com/stella/folio/pull/2871) [`3514b10`](https://github.com/stella/folio/commit/3514b103538d95edd72f06f92e1212061b235b9d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Canonicalize empty formatting on newly split paragraphs so later property edits remain exactly undoable.

- [#2668](https://github.com/stella/folio/pull/2668) [`f1646a4`](https://github.com/stella/folio/commit/f1646a4f0bde467a526ebc712a3babdb1e47e959) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a paragraph mark whose split left no piece of the record it cut on one side no longer keeps that seam on an insertion standing there. So a boundary kept from that insertion stays beside the next piece whether the run's revisions are rejected at once or one at a time.

- [#2681](https://github.com/stella/folio/pull/2681) [`76942b5`](https://github.com/stella/folio/commit/76942b5c20ca78ecf1d1e6f9ba07536e6ff3d888) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting changes that remove the pieces a paragraph split left at both of its edges now ends that split's cut there, so rejecting the split's mark joins the paragraphs as plain neighbours: an empty wrapper the author made beside the removed piece stays, and a boundary kept from an insertion at that edge stays beside the next record, whether the revisions are rejected at once or one at a time. A split seam with pending insertions standing at it is again kept on them when the record's other piece lies beyond a later paragraph, so the record joins again once they are rejected.

- [#2690](https://github.com/stella/folio/pull/2690) [`4bbbf3d`](https://github.com/stella/folio/commit/4bbbf3d4761becef4a0bfa87d22d95fc887aa08f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In a canonical room, an open header, footer or note editor keeps its caret and selection where its text went when a peer edits the story: the session holds the selection each such editor shows (`CanonicalSession.recordStorySelection` and `storySelection`) and moves it over every peer's edit by paragraph and offset, as it moves the body's. `transformPosition` maps a position past zero-width children, such as a caret just after a note's own mark, over a paragraph split or direct deletion that does not cut at its offset.

- [#2968](https://github.com/stella/folio/pull/2968) [`02e7789`](https://github.com/stella/folio/commit/02e7789e08e83b61dd892ab2f7c00117a13e478c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting a tracked replacement no longer gives the id of a record it struck whole to a piece it cut off a later record with the same author and date; the piece keeps its own id, as in the direct edit.

- [#2702](https://github.com/stella/folio/pull/2702) [`b8b58c4`](https://github.com/stella/folio/commit/b8b58c44a378fe1fbaf3f27f966ddd82a5dc8769) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested deletion of a table row that starts a vertical merge the next row continues leaves the merged cell's content unmarked, and accepting it hands the merge and its words to the next row, as deleting the row directly does. Rejecting a suggested row insertion that starts a merge hands the merge on the same way.

- [#2573](https://github.com/stella/folio/pull/2573) [`27e014c`](https://github.com/stella/folio/commit/27e014c94a142d69786fcc6b0e09a19cc0afa062) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested row, column or table deletion the tracked table records cannot express is refused under its own capability gap, and a row holding a nested table says so.

- [#3087](https://github.com/stella/folio/pull/3087) [`e39735c`](https://github.com/stella/folio/commit/e39735c6bb3cf9a882216a9fa6a0db47a25a3b3e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggestion deleting or replacing over its author's ordinary pending insertion records the deletion instead of removing the insertion, so rejecting the suggestion gives it back.

- [#2918](https://github.com/stella/folio/pull/2918) [`82cce07`](https://github.com/stella/folio/commit/82cce07cc870ea3da4a4e3b54a4e2e5945a266cc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a suggestion's tracked paragraph property change, join or split that would fold into a pending property change belonging to another author or suggestion. A paragraph holds one property change and a suggestion is rejected by its id, so taking the record over discarded the other change when the suggestion was rejected or saved as rejected.

- [#3494](https://github.com/stella/folio/pull/3494) [`8d63072`](https://github.com/stella/folio/commit/8d630727020d3ae63d998e8e62da4d0fb3226554) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give each tracked paragraph put before a table under review its own inserted mark, so separate suggestions there keep their own properties.

- [#3186](https://github.com/stella/folio/pull/3186) [`aca6020`](https://github.com/stella/folio/commit/aca60207ff6774b24de3b202a3cc0df23a880c3b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound paragraph locator index updates by sharing unchanged pages across document edits.

- [#2879](https://github.com/stella/folio/pull/2879) [`b521d67`](https://github.com/stella/folio/commit/b521d678e626062ca0ca5d64917229db5d37b291) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve table-relative insertion conflict ownership across cell paragraph aliases.

- [#3176](https://github.com/stella/folio/pull/3176) [`9485c98`](https://github.com/stella/folio/commit/9485c98b475b11047357b25860293ba3a9824f36) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Table operations address tables in headers, footers and notes, so table, row and cell revisions there resolve.

- [#3277](https://github.com/stella/folio/pull/3277) [`da3d107`](https://github.com/stella/folio/commit/da3d1077f802553afd4f0ba524a06e235ca6e161) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Row and column edits, and resolving row and cell changes, work on a table holding bookmark markers beside its rows or a row's cells: a marker keeps its place among the rows or cells that remain (a start before the one that followed it, an end after the one that preceded it), and each edit's inverse restores it exactly. Row edits on such a table were refused, and removed cells left the markers beside them out of range.

- [#3300](https://github.com/stella/folio/pull/3300) [`679773b`](https://github.com/stella/folio/commit/679773bc44e1288831631f5669fe8373789b7676) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested row insertion or deletion is given the revision ids its row's cells record, rather than those its anchor paragraph would need.

- [#3451](https://github.com/stella/folio/pull/3451) [`4a6bd8f`](https://github.com/stella/folio/commit/4a6bd8f0e5ddaa8d2e539d5572d1f3de06f9ed0a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - End a text box whose content ends in a table with a paragraph on the canonical session; an unedited save keeps the box as read.

- [#3301](https://github.com/stella/folio/pull/3301) [`0457292`](https://github.com/stella/folio/commit/0457292855c656544db7df7628f8941485dd59e9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Table operations on a table inside a text box get the revision ids they need, and an operation compiled in the box's own story passes the table operation checks.

- [#3231](https://github.com/stella/folio/pull/3231) [`ff440a2`](https://github.com/stella/folio/commit/ff440a29404bf722bd59662f7241c3f2c1574c0f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Check supplied tracked-edit IDs against the package identity index.

- [#2779](https://github.com/stella/folio/pull/2779) [`1aca979`](https://github.com/stella/folio/commit/1aca97908a45740dc9e09f4b90a00197bc3052d6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested column insertion or deletion applies to a table whose file states no column widths. The suggestion records the grid as the file states it, its columns without a measure, and rejecting it gives the table back with no widths.

- [#3263](https://github.com/stella/folio/pull/3263) [`36b6da5`](https://github.com/stella/folio/commit/36b6da58a30cc1ac6a8ccdfecf96f19d2e64d590) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record a tracked replacement of exactly a content control's text inside the control, so the control stays when the change is accepted or rejected; the canonical session no longer refuses a suggestion typed over a control's whole text.

- [#3064](https://github.com/stella/folio/pull/3064) [`774d0ba`](https://github.com/stella/folio/commit/774d0ba7069fa929c22a6d22e222a57d570aa393) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of an empty paragraph alone in its container applies again, recording nothing, as its direct deletion changes nothing.

- [#3041](https://github.com/stella/folio/pull/3041) [`1f049f5`](https://github.com/stella/folio/commit/1f049f595fb51a0a185c445bf3f634ab78a2fb5a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of an empty paragraph whose mark a pending deletion already took is refused: it would record nothing.

- [#2998](https://github.com/stella/folio/pull/2998) [`2496e7c`](https://github.com/stella/folio/commit/2496e7c9a0d07e239d3e0f9d97fa89012a33f55f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Track the deletion of a paragraph whose mark a pending deletion already took, or whose content holds tracked changes: the earlier deletion record stands, and the content is deleted as a tracked range deletion over it is.

- [#2678](https://github.com/stella/folio/pull/2678) [`ef78144`](https://github.com/stella/folio/commit/ef7814430fdb7733f54f6308fde06a0669b78a9f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked deletion of a range takes in a record it covers whole that holds a comment anchor, such as a content control or hyperlink holding a comment's end. Accepting the deletion removes the anchors it holds, as deleting the range directly does, so the two give the same document; rejecting it gives the document back. A bare comment anchor is still left standing by a planned deletion, which deletes around it.

- [#2984](https://github.com/stella/folio/pull/2984) [`e643036`](https://github.com/stella/folio/commit/e643036e7684bfa0079c193fbcdbabeb0225ffc7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accept a suggested insertion at the point where a concurrent insertion cut a tracked change or content control, in the authority's order, instead of rejecting it.

- [#3124](https://github.com/stella/folio/pull/3124) [`969c866`](https://github.com/stella/folio/commit/969c8668565b5b8c035746f179456184df1a7ad0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use indexed identity lookups for tracked paragraph insertions.

- [#2699](https://github.com/stella/folio/pull/2699) [`c7ce21e`](https://github.com/stella/folio/commit/c7ce21e1f722a93bcfe94778089f836455ba634b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Setting, removing and inserting a hyperlink, and generating a table of contents, are recorded as suggestions in the canonical session. A link change strikes each stretch of the changed text and inserts its changed copy after it, so rejecting brings the text back as it was and accepting leaves what the edit made directly leaves; a table of contents records each paragraph it inserts and the split it makes. A suggested link change that would leave a bookmark or comment marker alone inside a piece of a link, or copy a drawing, shape or note reference, is refused under the `linkSuggestionShape` rule (reason `unrecordableLinkSuggestion`), with the `editor.linkSuggestionUnavailable` message in every catalog language; the same change made directly applies.

- [#2695](https://github.com/stella/folio/pull/2695) [`d19c09f`](https://github.com/stella/folio/commit/d19c09f87ade2458f37369d9b67c96598518c47c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Concurrent tracked changes of one paragraph that meet are accepted rather than refused: a cut at a tracked change's edge or inside its range, a direct deletion reaching one of its edges or covering it, and overlapping tracked deletions, where the deletion sequenced first holds the shared content. The authority's order decides the document and every peer replays to it; `trackedOrderPair` names these pair kinds. A peer that applied its own edit of such a pair first gets the document as a `replaced` change, which now carries `moves`, the rebased change, when the two documents hold the same paragraphs at the same lengths. A canonical session moves its selection over `moves` instead of falling back to the nearest position.

- [#3379](https://github.com/stella/folio/pull/3379) [`d12d0d2`](https://github.com/stella/folio/commit/d12d0d2442c530e98280647b8af9c126a4dd8130) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a tracked paragraph inserted after a table that ends a content control inside the control; rejecting it removes it whole.

- [#3266](https://github.com/stella/folio/pull/3266) [`21095d0`](https://github.com/stella/folio/commit/21095d028b9123c201c9307446b47d551190162a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Track the deletion of a paragraph that holds a content control: the deletion holds the control whole, accepting removes it and rejecting gives it back.

- [#2727](https://github.com/stella/folio/pull/2727) [`d1e7461`](https://github.com/stella/folio/commit/d1e7461ab5ca3ebc4d0b794ff07a84c16aa1ebf2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested row or table deletion records rows that hold a nested table, bookmark or comment boundaries, or a content control around a cell: the words inside are marked as the row's are, and accepting or rejecting resolves them with the row. A suggested table deletion over a row already suggested for deletion marks every other row and leaves that row's own suggestion as it is.

- [#3292](https://github.com/stella/folio/pull/3292) [`4651974`](https://github.com/stella/folio/commit/46519749e6e40b795b9b96a85e5ec612cd9bb842) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Tracked deletions of a container's last paragraphs accept as the same deletions made directly in any order. Deleting a paragraph whose break a pending deletion already takes, into paragraphs that keep nothing once accepted up to the container's end, now also records the break before it deleted, as when that paragraph is last; accepting no longer leaves an empty paragraph behind.

- [#2742](https://github.com/stella/folio/pull/2742) [`df8876f`](https://github.com/stella/folio/commit/df8876f99e69e84e1e63df7a582da8d2a9c793cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested split of a vertical merge records each row's cell leaving the merge, so rejecting the suggestion restores the merge.

- [#2815](https://github.com/stella/folio/pull/2815) [`d9143b5`](https://github.com/stella/folio/commit/d9143b5c4c9a0e6f56ac5e13fcec9d4ce71bffc8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Typing, inline content insertions and inline record splits and joins in a paragraph pass a concurrent paragraph insertion before or after that paragraph, direct or tracked, in either sequence order, rather than being refused.

- [#2980](https://github.com/stella/folio/pull/2980) [`cf54654`](https://github.com/stella/folio/commit/cf5465452ac71756625fdbb34826c7e9d20af48f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Let an undo of typing pass a concurrent split of its paragraph away from the undone text, or a join of its paragraph, instead of rejecting it.

- [#2796](https://github.com/stella/folio/pull/2796) [`4406db9`](https://github.com/stella/folio/commit/4406db9fd81466e463219fd1e188166f11fed4d2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Undo of a committed document operation on the canonical session restores the exact prior document when the operation removed notes with their references and list rendering then renumbered the paragraphs after it: `inverseMeetingDocument` rebases each whole-part restore of an inverse onto the document it meets when the inverse runs. A rejected undo of a document-operation batch names its cause: `documentChanged` when the document changed after the batch, `storyUnavailable` when the story the batch edited is not open, `inverseRefused` (with the session's message) when the batch's inverse cannot be staged, and `documentNotEditable` when the host does not take the undone document.

- [#2466](https://github.com/stella/folio/pull/2466) [`7350f0b`](https://github.com/stella/folio/commit/7350f0b807ecd8252a11324e8db4fe5384539aa8) Thanks [@dependabot](https://github.com/apps/dependabot)! - Script detection uses the Unicode 18 Script and Script_Extensions tables.

- [#3144](https://github.com/stella/folio/pull/3144) [`9933560`](https://github.com/stella/folio/commit/993356051bbe53ee2f8c090adb5a5a57f4ffc3d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse immutable paragraph splice boundaries when checking changed blocks and updating identity indexes.

- [#3233](https://github.com/stella/folio/pull/3233) [`fb689e8`](https://github.com/stella/folio/commit/fb689e81c58738becc5bca0d5670cb34f24e96d5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Validate the added paragraph of a tracked insertion when its checked neighbor is retained unchanged.

- [#3150](https://github.com/stella/folio/pull/3150) [`fe4a82a`](https://github.com/stella/folio/commit/fe4a82acfd943abaad95428b3983b625ae0c8949) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Validate eligible tracked paragraph insertions before materializing the final document once.

- [#2930](https://github.com/stella/folio/pull/2930) [`aac3709`](https://github.com/stella/folio/commit/aac370985958522b3e6075882c947c2b71fd2cb1) Thanks [@dependabot](https://github.com/apps/dependabot)! - The bundled WebAssembly modules are regenerated with wasm-bindgen 0.2.130.

- [#2938](https://github.com/stella/folio/pull/2938) [`a2b285f`](https://github.com/stella/folio/commit/a2b285f1236cefbc31e7024eb923aa83c59ae7d5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A replacement over whole paragraphs that hold pending insertions, words or breaks, retires the comments they cover, as it does over paragraphs without them: a direct edit removes any author's insertion with its paragraph, and a suggestion retracts its own author's. In a canonical session, typing over everything after a suggested paste beside a comment applies instead of being refused.

- [#3117](https://github.com/stella/folio/pull/3117) [`a0b7452`](https://github.com/stella/folio/commit/a0b7452d3e1f45b64be6dc04404412b193680acd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The optimistic sequencing client withholds a submitted batch that no longer rebases until the authority answers it: a rejection drops it with a notice, and a sequenced one comes back as a document change instead of being reported dropped and then applied unseen.

## 0.43.1

### Patch Changes

- [#2449](https://github.com/stella/folio/pull/2449) [`348cc23`](https://github.com/stella/folio/commit/348cc23ab966f5b9cf6334305648140a11a1c898) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Tracked block moves now apply in the experimental canonical session. A `deleteBlock` and an `insertAfterBlock` or `insertBeforeBlock` that share a `moveId` in one batch are recorded as `w:moveFrom` and `w:moveTo`, paired as the default engine pairs them, and an unpaired id is reported as `unpairedMove`. `INSERT_BLOCKS` and `DELETE_BLOCKS` take `move`. A block deletion the following paragraph survives now removes the paragraph by its identity.

- [#2450](https://github.com/stella/folio/pull/2450) [`46dda37`](https://github.com/stella/folio/commit/46dda37b0c11ad2cda0fed859474c64291771c4f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A `newList` numbering request on `insertAfterBlock`, `insertBeforeBlock` or `setBlockParagraphProperties` now applies in the experimental canonical session. It mints a list instance per operation and kind, as the default engine does. The `insertParagraphs` and `formatParagraph` editor intents take the new instance and create it first. `NewListInstance` is exported from `@stll/docx-core/ops`.

- [#2479](https://github.com/stella/folio/pull/2479) [`892aafa`](https://github.com/stella/folio/commit/892aafa4b63d6e8f5cc29980311f55ac1d95fb6c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Tracked paragraph insertions and deletions next to tables, block-level range boundaries and a container's final paragraph now resolve one record at a time without reminting the anchor's record ids, and a tracked deletion before a table that cannot take its mark is refused up front. Direct joins and paragraph deletions now leave the caret in a paragraph the edit keeps.

- [#2454](https://github.com/stella/folio/pull/2454) [`bcee67a`](https://github.com/stella/folio/commit/bcee67afd150cbc4a5938363d40c5f67eb874cea) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A commented public operation on a paragraph inside a table cell or a block content control now applies in the experimental canonical session. It no longer refuses with "The comment anchor no longer exists." A commented insertion beside such a paragraph now carries its comment too. `storyParagraphs` is exported from `@stll/docx-core/ops`.

- [#2527](https://github.com/stella/folio/pull/2527) [`17f7f28`](https://github.com/stella/folio/commit/17f7f28c2d1ef0db4bad6d7cad4dde18f801759b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Typing in the canonical session applies each keystroke's edit once: the compiled edit hands back what it applied, and the session no longer applies it again.

- [#2447](https://github.com/stella/folio/pull/2447) [`c850dcd`](https://github.com/stella/folio/commit/c850dcd18b4aed0a121a130fed21cb9c54ae771f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked `deleteBlock` of a paragraph that a table follows now applies in the experimental canonical session, as the default engine records it: its words and its own mark are marked deleted. `DELETE_BLOCKS` records this with a revision when the block list still ends with a paragraph. A new `deleteParagraphs` editor intent compiles to it.

- [#2479](https://github.com/stella/folio/pull/2479) [`892aafa`](https://github.com/stella/folio/commit/892aafa4b63d6e8f5cc29980311f55ac1d95fb6c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Canonical public operations: an inline insertion reads CRLF and CR as one line break; an insertion beside a paragraph in a table cell goes beside the table; an insertion beside a paragraph the same batch deletes applies when a block follows that paragraph; a move whose other half does not apply saves as a plain insertion or deletion, reported unpaired; an undone insertion's paragraph id is not reused; and an open header, footer or note editor keeps its caret through an operation on its story. A list reference that names no list stays out of lists started later.

- [#2470](https://github.com/stella/folio/pull/2470) [`217941f`](https://github.com/stella/folio/commit/217941f86dddaf7b6110e315cf3f85bdde41e659) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In the experimental canonical session, a tracked replacement now reads as a redline: the struck text comes first and the new text follows it, as the default editor saves it. This holds for typing, pasting and pressing Enter over a selection, and for public replacements. Text replacing a selection that ends inside a content control stays inside that control instead of splitting it, and replacing exactly the words a comment covers keeps the comment on the new text.

- [#2435](https://github.com/stella/folio/pull/2435) [`95488ec`](https://github.com/stella/folio/commit/95488ec05c97a9b4769c3e1f5880fc9840d9b489) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The experimental canonical session now opens documents whose text holds a literal tab character, and documents with a header, footer or note that holds no blocks. The session reads a literal tab as a tab element and gives an empty story an empty paragraph, in its own model only. A part whose story is unedited saves exactly as it was read; a story the user edits saves in the session's spelling. `replaceStoryBody` is exported from `@stll/docx-core/ops`.

- [#2504](https://github.com/stella/folio/pull/2504) [`97b5ca5`](https://github.com/stella/folio/commit/97b5ca558476d93f6ea0f8ff6aed1e0d5b155f3c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked batch that deletes the paragraph ending the body and inserts beside it now applies both operations in the experimental canonical session, as the default engine does. `DELETE_BLOCKS` with a revision now removes a paragraph mark that the same author's pending insertion added, joining that paragraph to the next one, instead of refusing the deletion.

- [#2475](https://github.com/stella/folio/pull/2475) [`05bfbdf`](https://github.com/stella/folio/commit/05bfbdf74f2e37c31a181b710f87c063d4d5c11b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In the experimental canonical session, consecutive tracked typing and deletion by one author now read as one change, as in the default editor. A word typed while suggesting saves as one insertion instead of one per character, and a run of Backspace or Delete saves as one deletion. Text typed beside the author's own pending insertion goes into it, and a deletion beside the author's own pending deletion continues it, also when that change came with the file. Another author's change is never continued. Undo still takes back each edit, and rejecting gives the original text back with its runs whole.

- [#2493](https://github.com/stella/folio/pull/2493) [`963c5f6`](https://github.com/stella/folio/commit/963c5f69c7296f279da9288d67c25e9d8780c9d6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Checking a document's comments after an edit reads only the paragraphs the edit replaced: each paragraph's comment anchors, text edges and text boxes are read once per paragraph.

- [#2489](https://github.com/stella/folio/pull/2489) [`76afe3a`](https://github.com/stella/folio/commit/76afe3adcf8a95bbc51a53f1a8d24ff5c3c4466b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The comment anchor check reads each comment's anchors from one pass over the document's anchors instead of filtering all of them once per comment, so documents with many comments check in time proportional to their anchors.

- [#2522](https://github.com/stella/folio/pull/2522) [`02b6a24`](https://github.com/stella/folio/commit/02b6a24af33b00fe78c4b543a6f954f14cd77af9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Replacing a range directly whose deletion keeps a bookmark or another range boundary inside a wrapper or inline content control now puts the new text in that wrapper, after the boundary, instead of splitting the wrapper in two. The direct edit now gives what accepting the same edit made with tracked changes gives.

- [#2472](https://github.com/stella/folio/pull/2472) [`32e7052`](https://github.com/stella/folio/commit/32e705214b247158125aca904eefd2a519802ba4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `@stll/docx-core` and `@stll/folio-core` no longer depend on `@stll/docx-utils`. The OOXML namespace table they read from it, `OOXML_NS` (with its `OoxmlPrefix` type), now comes from `@stll/docx-core/xml/namespaces`, with the same entries. Installing folio no longer brings in `@stll/docx-utils` or `jszip`.

- [#2495](https://github.com/stella/folio/pull/2495) [`e205e8f`](https://github.com/stella/folio/commit/e205e8fe07e7ba3cc1ddee2fb592550a447ea7e9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An edit no longer counts every revision and content-control id of the package to check the ids it creates: the checks ask the package's id index for the ids in question.

- [#2528](https://github.com/stella/folio/pull/2528) [`466d3b7`](https://github.com/stella/folio/commit/466d3b7ef9a35a71bb52ce2bbef6a976acc7404e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Edits look up revision and content-control ids without counting the whole document again when the previous version they were counted from is gone.

- [#2490](https://github.com/stella/folio/pull/2490) [`eb1ca54`](https://github.com/stella/folio/commit/eb1ca54500e47d62812fbfd2ee74051441e92bd9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Adds a per-package index of revision and content-control ids that is derived from the previously indexed package by walking only the records an edit replaced.

- [#2532](https://github.com/stella/folio/pull/2532) [`9fad4c9`](https://github.com/stella/folio/commit/9fad4c97e670c414c111f8a5d690c86511bbc75f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Applying an operation decides whether it removed a note's last reference from the top-level blocks the operation changed, instead of reading every paragraph of the document before and after.

- [#2477](https://github.com/stella/folio/pull/2477) [`0b3a51b`](https://github.com/stella/folio/commit/0b3a51b7f73f7dd624bd566df828feb05d9f3709) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting only the inserted paragraph mark of a paragraph that keeps its identity, after paragraphs were inserted at the end of a container, now keeps the next paragraph's pending mark and pending property changes instead of dropping them.

- [#2516](https://github.com/stella/folio/pull/2516) [`bb6a653`](https://github.com/stella/folio/commit/bb6a653506dc635112ac0b21647e43bd98d191c2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a tracked or suggested paragraph inserted after a paragraph that a block content control (or any block other than a paragraph or a table) follows no longer leaves an empty paragraph behind. As at the end of a container, the anchor's mark is recorded as the inserted one, so rejecting it joins the two paragraphs back into the anchor.

- [#2492](https://github.com/stella/folio/pull/2492) [`0df0787`](https://github.com/stella/folio/commit/0df07878224969d3a311d25e6cd42506cad9687a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In the canonical session, typing over a selection right beside your own pending deletion now joins the struck text to that deletion, as the default editor records it, instead of saving a second deletion beside it.

- [#2532](https://github.com/stella/folio/pull/2532) [`9fad4c9`](https://github.com/stella/folio/commit/9fad4c97e670c414c111f8a5d690c86511bbc75f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Applying an operation finds the paragraphs whose authored property snapshots it carries over among the top-level blocks the operation changed, instead of reading every paragraph of each changed story twice.

- [#2460](https://github.com/stella/folio/pull/2460) [`11a46ec`](https://github.com/stella/folio/commit/11a46ec10cc10c2ecd43ca0eda5d8f731dc9a3fa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In the canonical session, suggested batches can now carry `commentOnRange` and `commentOnBlock`, and a comment attached to any suggested operation now belongs to its suggestion. A proposed comment is listed with its suggestion (a new `"comment"` kind in `getSuggestions`; a suggestion that only proposes a comment applies as `"direct"`), is left out of saved files while pending, travels through the pending-suggestion export and restore, is removed when the suggestion is rejected and stays as an ordinary comment when it is accepted. A tracked or suggested edit's attached comment now covers the text the edit changed, as in the default engine, instead of the whole replaced range.

- [#2457](https://github.com/stella/folio/pull/2457) [`f42e5be`](https://github.com/stella/folio/commit/f42e5be33332b5cb4dc9fe771df0489ed03aaf71) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In the experimental canonical session, suggested batches can now carry `splitBlock` and `mergeBlockWithNext`. Each records tracked changes stamped with the suggestion: rejecting gives the source back, and accepting adopts them. A canonical public split now keeps the paragraph's identity on the first half, as the default engine's split does. Rejecting a canonical split's inserted mark gives the paragraph back under its own id.

- [#2521](https://github.com/stella/folio/pull/2521) [`db7d15f`](https://github.com/stella/folio/commit/db7d15f057a7926a7cac4ec861da8a3413ca077c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked or suggested deletion of a paragraph that a block content control, custom XML or a preserved block follows now applies in the canonical session instead of being refused. The paragraph's words and its own mark are recorded as deleted; accepting removes the paragraph and rejecting gives it back. Accepting a deleted mark there removes a paragraph left with nothing in it, keeps one that still holds content, and runs content that accepting will take away into the content control's first paragraph, so resolving the words and the mark together or one at a time, in either order, leaves the same document.

- [#2526](https://github.com/stella/folio/pull/2526) [`257435a`](https://github.com/stella/folio/commit/257435a14c1e6038d43022f51346d3bd3ea16b0b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Typing over a selection beside or across the author's own earlier tracked deletions applies whatever times those deletions carry; the struck words and the typed text keep distinct revision ids.

- [#2502](https://github.com/stella/folio/pull/2502) [`ef02ec4`](https://github.com/stella/folio/commit/ef02ec44e62efa8fdcf145a7bad4f5c23267dcc8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Checking which new ids an edit really added now reads the package's id index instead of collecting every id in the package.

## 0.43.0

### Minor Changes

- [#2404](https://github.com/stella/folio/pull/2404) [`05285fc`](https://github.com/stella/folio/commit/05285fc15612b3a6c88787ec61b1591ad0a47a73) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A revision can name the pending suggestion it belongs to: `TrackedChangeInfo.suggestionId`, set from a new `suggestionId` on an operation's revision stamp. A record of one suggestion never merges into a record of another suggestion or into an ordinary tracked change. The id is held in memory only. `withoutPendingSuggestions` returns a document with every pending suggestion rejected, and the experimental canonical session's save writes the file through it, so a pending suggestion never reaches the file as an ordinary tracked change. If a suggestion cannot be rejected, the save refuses instead of writing it.

- [#2414](https://github.com/stella/folio/pull/2414) [`bb3018d`](https://github.com/stella/folio/commit/bb3018d5a43926c0ba2d8760e56a05724089b823) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The experimental canonical session supports pending suggestions. A document operation batch in suggested mode records tracked changes that carry the operation's suggestion id, never folding into a suggestion already pending. The review surface lists them as suggestions, including formatting suggestions. Accepting one adopts it as ordinary tracked changes by the accepting author; rejecting one rejects its revisions. Both are journalled, so undo restores the suggestion, and the React and Vue editors route accept and reject to the session. The engine gains `setRevisionInfo`, which rewrites revision metadata with an exact inverse, `suggestionRevisionInfos` and `pendingSuggestionIdsIn`.

- [#2403](https://github.com/stella/folio/pull/2403) [`9ccf3b0`](https://github.com/stella/folio/commit/9ccf3b0f66c6a8524c495e7029b3e2c55191cd34) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A complex field whose begin and end both sit inside a hyperlink is now one child of the link (`Hyperlink["children"]` admits `ComplexField`). A table of contents entry is one link over the heading text, a tab and a PAGEREF page number; the page number used to be read as loose field-character runs. The editor dropped those runs on save, so an edited entry lost its page number field, and the experimental canonical session refused the document. The field now saves inside the link it was read from, the editor round trip keeps it there, and the canonical session opens such documents. Applying a link over a selection that holds a field still leaves the field between the linked runs.

- [#2458](https://github.com/stella/folio/pull/2458) [`a5228eb`](https://github.com/stella/folio/commit/a5228eb9b501605d5e919898e12d2abed531674f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A document whose review provenance folio cannot read now opens instead of being refused. Review provenance is the `folio:` attributes in the ignorable review-history namespace, and the cases are another version, an unknown field, or a malformed value. The reader drops that provenance, reads the revisions it annotates as they are, and reports each dropped value as a `review-provenance-dropped` parse warning. What folio writes is still validated strictly.

### Patch Changes

- [#2370](https://github.com/stella/folio/pull/2370) [`15a57b5`](https://github.com/stella/folio/commit/15a57b54cb48a7b9095fd2f7ba1540a2e00c5a5f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A comment range marker written between blocks, beside a content control for example, is kept as a typed block in the document model and both editors, and is saved back between the same two blocks. The comment used to read as a point comment, and a save after any edit wrote its start twice under the same id.

- [#2328](https://github.com/stella/folio/pull/2328) [`5912d35`](https://github.com/stella/folio/commit/5912d350a009e2430cc848239a79bde00d1cafd9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A hyperlink can now be added, changed or removed in a paragraph holding comment ranges, a comment reference, proofing marks or a permission range. The edit was refused without telling the user.

- [#2445](https://github.com/stella/folio/pull/2445) [`51443e1`](https://github.com/stella/folio/commit/51443e18a4e0163f5a53f84f32ac6a6d26fe62b0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Block-level public operations in the experimental canonical session now leave paragraph ids where the default engine leaves them, so host references stay valid across engines. Inserting after the last paragraph of a list and deleting that paragraph directly now apply instead of refusing. In the engine, `INSERT_BLOCKS` now records a tracked insertion after a paragraph that no paragraph follows, and a tracked insertion before a table. A new `insertParagraphs` editor intent compiles to `INSERT_BLOCKS`, and `joinParagraphs` takes a `survivor` for direct joins. A paragraph mark gains an in-memory `survivor` that is never written to a file. Rejecting a suggested insertion at the end of a list gives the document back exactly.

- [#2382](https://github.com/stella/folio/pull/2382) [`b305fdf`](https://github.com/stella/folio/commit/b305fdf59b85d5145516d623fa4944bd6c08c7dd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - In the experimental canonical session, text typed over a selection that holds comment, bookmark or permission markers now stays together and in order. The first character no longer lands on one side of a marker and the rest on the other, and in suggesting mode the typed run no longer splits around the struck text.

- [#2394](https://github.com/stella/folio/pull/2394) [`fa3b0bf`](https://github.com/stella/folio/commit/fa3b0bfd7e594be9d272510c075411dcb94c236a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The experimental canonical session now opens a document with inline content controls. Typing inside a control keeps the text in it, and a caret just inside a control's edge types into the control while one just outside types beside it, in editing and suggesting mode. Deleting inside a control keeps the control, and locks hold for inline controls as for block ones. Replacing a selection of text inside a control refuses for now, since it would cut the control in two.

- [#2395](https://github.com/stella/folio/pull/2395) [`06c4073`](https://github.com/stella/folio/commit/06c4073d886ae0e8d1eb569adcf6c3f82aa97d80) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The experimental canonical session now runs the editor's insert table and delete table commands, in editing and suggesting mode. An inserted table comes with the empty paragraph the editor places after it; in suggesting mode its rows and paragraphs are recorded as inserted, and a deleted table's rows as deleted. The shared intent compiler gains a `tableBlock` intent whose inserted table takes fresh paragraph identities.

- [#2430](https://github.com/stella/folio/pull/2430) [`544fb6d`](https://github.com/stella/folio/commit/544fb6d9933a52f9ad67e2262966297dc814c1be) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Inserting a table in the experimental canonical session now works from a selection of the whole document and from selected cells, placing the table where the editor does. A whole-document selection puts it at the start of the document, before the first paragraph. Selected cells put it after their table, with an empty paragraph on each side so the two tables never touch. The table block intent gains `side` and `separated`.

- [#2390](https://github.com/stella/folio/pull/2390) [`d4c8b0d`](https://github.com/stella/folio/commit/d4c8b0d09ac84da1984831a3dd026735033bcebb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The experimental canonical session now runs the editor's row commands: insert a row above or below and delete rows, in editing and suggesting mode. Each saves what the editor would save and undoes exactly; deleting every row removes the table. The shared intent compiler gains a `tableRow` intent whose inserted row takes fresh paragraph identities.

- [#2378](https://github.com/stella/folio/pull/2378) [`db7686a`](https://github.com/stella/folio/commit/db7686af8ddee78affcfa9143795e813301664e4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A comment whose covered text is all deleted is now removed on save also when a bookmark or another range boundary sat inside its range. Such a boundary marks a position, not content, so it no longer keeps an empty comment alive. A source comment whose range holds nothing but such a boundary still opens as a point comment and is kept.

- [#2393](https://github.com/stella/folio/pull/2393) [`4ed62d6`](https://github.com/stella/folio/commit/4ed62d651d781ccc7fb119aee605a9e3f8bed07f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Editing a paragraph whose tracked insertion or deletion holds a point comment reference (a `w:commentReference` with no comment range) no longer fails: splitting, inserting into and deleting from that change keep the reference inside it.

- [#2421](https://github.com/stella/folio/pull/2421) [`33c708f`](https://github.com/stella/folio/commit/33c708f9dbc455a60a6fb749e180527bf99a2f5b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `generateRedlineDocx` aligns every story of one comparison against a single alignment allowance, as `compareDocxVersions` does. The reference sequencer replays a stored outcome only to the actor that first used its operation id, and the optimistic client refuses a broadcast that reuses a pending operation id under another actor.

- [#2375](https://github.com/stella/folio/pull/2375) [`e53e82a`](https://github.com/stella/folio/commit/e53e82adef1c1a2b99a050c255369ae92d4b423e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Document operations now honour a content control's `w:lock` in every mode. Under `contentLocked` or `sdtContentLocked`, an operation that changes text, formatting or blocks inside the control is skipped. Under `sdtLocked` or `sdtContentLocked`, an operation that removes the control is skipped, including deleting the paragraph that holds it. Either way the operation reports the new reason `contentControlLock` and changes nothing. A document comparison still shows changes inside locked controls.

  `@stll/docx-core/model` exports the rule: `contentControlLockAllows(lock, change)` decides whether a lock allows editing a control's content or removing the control, and `lockedControlViolation` applies that decision to the stretch a change touches.

- [#2405](https://github.com/stella/folio/pull/2405) [`c8e91a8`](https://github.com/stella/folio/commit/c8e91a8d96a548becc38dbf77fc130ad734dd8bb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving a tracked replacement's deletion and insertion one at a time, in any order and with either decision, gives the same runs as resolving them together, at any depth. Replacing part of the text inside an inline content control in the canonical editor now keeps the new text in the control, in editing and suggesting mode.

- [#2429](https://github.com/stella/folio/pull/2429) [`2d9d34c`](https://github.com/stella/folio/commit/2d9d34cf49b4912926e9fe51143d31237319fa28) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Setting, removing or inserting a hyperlink now works over a run with a pending formatting change, and over tracked insertions, deletions and moves. Before, the experimental canonical session refused these edits. A record the link edit cuts keeps its id on the first half, and the halves after it take fresh ids, as any cut does. Text inside a tracked change is linked inside that change, which is how the file nests it.

- [#2442](https://github.com/stella/folio/pull/2442) [`682b38b`](https://github.com/stella/folio/commit/682b38b917ab295dc3ece1540d1affd242b0aee5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A canonical session normalizes list rendering after an edit only from the last checkpoint before the first paragraph the edit changed, and only until the list counters match the earlier run again, instead of over every paragraph of the story. `paragraphPropertySnapshotToken` names the snapshot registration a paragraph carries, so a pass can tell that registering again would change nothing.

- [#2444](https://github.com/stella/folio/pull/2444) [`bed4019`](https://github.com/stella/folio/commit/bed401918d79be1d2faaa4101b97a91788e3db8b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The id census an editor intent reads now counts a new version of the main story by the blocks an edit replaced also when the story holds tables, instead of walking the whole package again.

- [#2432](https://github.com/stella/folio/pull/2432) [`130e31f`](https://github.com/stella/folio/commit/130e31f98c26b43190666a993be19698b003b39d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An edit reads which notes the document still references once per paragraph it replaced, instead of reading every paragraph of both versions.

- [#2438](https://github.com/stella/folio/pull/2438) [`780dbc1`](https://github.com/stella/folio/commit/780dbc17c0bb3c8abb1c8b54515f152f4a10c2a2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Internal cleanup for the stricter shared lint rules: merged duplicate imports, block-scoped `case` bodies and `flatMap` in place of `map().flat()`. No behavior change.

- [#2428](https://github.com/stella/folio/pull/2428) [`eb49cc4`](https://github.com/stella/folio/commit/eb49cc4ee1ca8f286636ef69fc0fcdd412aacfaf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Planning a tracked edit reads a package's revision, content-control and reserved ids once per package instead of on every attempt, so a tracked replacement in a long document plans about three times faster.

- [#2463](https://github.com/stella/folio/pull/2463) [`81df722`](https://github.com/stella/folio/commit/81df722735d08239662ab3e03fccab3f33c7350a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked insertion that records which paragraph rejecting its paragraph mark keeps (a block inserted after the paragraph ending its block list) now saves that record as its own ignorable review-provenance attribute, `folio:rejectionSurvivor`, and reads it back. Rejecting the insertion after a save and reopen now keeps the same paragraph, under the same id, as rejecting it before the save. A reader that does not know the attribute drops only the marker and joins the paragraphs as for any rejected mark.

- [#2448](https://github.com/stella/folio/pull/2448) [`4acfd6a`](https://github.com/stella/folio/commit/4acfd6a850ec550af577cb9defcb2944952ccb85) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Planning a tracked edit applies each piece once when it takes no new ids, instead of first counting the ids it needs, so tracked typing in long documents spends less time planning.

- [#2431](https://github.com/stella/folio/pull/2431) [`677ef35`](https://github.com/stella/folio/commit/677ef35e8560e8ee3c54ac047ff6ba3c538d63de) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The id census reuses what each unchanged record contributes, so a document an edit made is counted by walking only the records the edit replaced.

- [#2401](https://github.com/stella/folio/pull/2401) [`2919a0d`](https://github.com/stella/folio/commit/2919a0d87922ba84ffa04e8da939101451bc00c6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A table row or cell inside a content control no longer counts that control twice. The parser records such a control in the row's or cell's control stack and again in its wrapper order, and the identity check took the two for two controls carrying one id, so the experimental canonical session refused the document.

- [#2347](https://github.com/stella/folio/pull/2347) [`4cb7cb7`](https://github.com/stella/folio/commit/4cb7cb7d740a588171a46d0d90104b3b31bf4884) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A text box keeps everything it holds. Content controls, custom XML, bookmarks, `w:altChunk` and other block markup in `w:txbxContent` used to be left out of the model and were dropped from the saved file when the paragraph hosting the box was edited, along with any text in them. Text in a control or custom XML now shows in the box, and an `w:altChunk` shows its marker.

- [#2391](https://github.com/stella/folio/pull/2391) [`98682ef`](https://github.com/stella/folio/commit/98682ef81045667104aedb888e1aa20964ba4459) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A suggested column deletion now keeps the table grid while it is pending and marks each deleted cell (`w:cellDel`), as the editor's own command does; accepting removes the column with its grid entry and rejecting restores the table exactly. The canonical session tracks a suggested column deletion instead of refusing it, and resolving revisions in a document that opens with a table no longer refuses.

- [#2372](https://github.com/stella/folio/pull/2372) [`0f4712a`](https://github.com/stella/folio/commit/0f4712a8a585435c1200ec4bccf892ddee4ea9c5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A table, row or cell with a pending tracked formatting change can now take another formatting change. In suggesting mode the change joins the pending record, which still holds the formatting from before both, so rejecting restores the original and accepting keeps the latest. An untracked change keeps the pending record. Before, any further change to that element was refused until the pending one was accepted or rejected.

## 0.42.1

### Patch Changes

- [#2333](https://github.com/stella/folio/pull/2333) [`0577924`](https://github.com/stella/folio/commit/0577924790bb43c2c588c451f632a846db5110a5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Applying a hyperlink over a selection holding a field no longer removes the field and its result on save. The runs around the field are linked and the field stands between them.

- [#2337](https://github.com/stella/folio/pull/2337) [`80c4d27`](https://github.com/stella/folio/commit/80c4d279d8e7b4716936bb102bfdae147e8f031d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A document with a table row holding no cells, which the file format allows, now opens instead of failing with an invalid document model. Row and column edits elsewhere in the table leave such a row where it is.

## 0.42.0

### Minor Changes

- [#2227](https://github.com/stella/folio/pull/2227) [`9cd0d03`](https://github.com/stella/folio/commit/9cd0d03924319bbfa1f16f3bd7821b1eb7e367f4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A batch now equals its operations applied one at a time with every operation that addresses a block before any operation that deletes or merges that block away. `@stll/docx-core/ops` exports the shared order (`orderAddressingBeforeRemoval`) and the rule for which deletions take a block out (`deletionRemovesBlock`). In tracked changes, properties a batch sets on a paragraph it merges away, or on the last paragraph it deletes, no longer survive the carried properties: accepting gives what the same batch applied directly gives.

## 0.41.0

### Minor Changes

- [#2189](https://github.com/stella/folio/pull/2189) [`99bad6b`](https://github.com/stella/folio/commit/99bad6be1c13bb8c78ca845374c9c36fdf49f8a1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read XML with `@stll/docx-core/xml`, folio's single-pass reader, everywhere: DOCX parsing, package and schema validation, numeric identity checks, preserved-markup checks and the CLI's package checks. `readXml` states its conformance (`ooxml`, `well-formed` or `tolerant`), `wellFormedXmlIssue` replaces separate validation passes, and resource limits and namespace resolution move to `@stll/docx-core/xml/limits` and `@stll/docx-core/xml/namespaces`. Markup that is not well-formed (mismatched or unclosed tags, attributes without quoted values, a DOCTYPE) is refused rather than repaired, and validators also refuse `<` in attribute values, `]]>` in text, `--` in comments and a misplaced XML declaration. `parseXmlWithFastXmlParser` is removed, and fast-xml-parser is no longer a dependency.

## 0.40.0

### Minor Changes

- [#2186](https://github.com/stella/folio/pull/2186) [`28dab55`](https://github.com/stella/folio/commit/28dab5569647efe7907960de4dc6f47ed4976d9f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read and write DOCX packages with `@stll/docx-core/zip`, a bounded ZIP archive built on the platform's DEFLATE, and add `@stll/docx-core/zip/stream` for streamed archives. Reads check every declared size, ratio and name before inflating and stop each entry at its declared size; writes pass unchanged entries through as their original compressed bytes. `writeZipPart`, `writeZipEntry` and the `no-string-zip-entry` lint rule are removed, `RawDocxContent.originalZip` is a `ZipArchive`, and JSZip is no longer a dependency.

### Patch Changes

- Updated dependencies [[`28dab55`](https://github.com/stella/folio/commit/28dab5569647efe7907960de4dc6f47ed4976d9f)]:
  - @stll/utf8-boundaries@0.2.0

## 0.39.2

### Patch Changes

- [#2178](https://github.com/stella/folio/pull/2178) [`6f9ba85`](https://github.com/stella/folio/commit/6f9ba852f5a9015ab70c5fbd90465931e4f3d684) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refresh native search and Rust dependency inputs while preserving canonical WebAssembly outputs.

- [#2184](https://github.com/stella/folio/pull/2184) [`c42a79f`](https://github.com/stella/folio/commit/c42a79f531a346e0e437b685a2e93cd25e48d7b4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rebuild canonical WebAssembly artifacts with Rust 1.99.1 and record the updated compiler input provenance.

## 0.39.1

### Patch Changes

- [#2161](https://github.com/stella/folio/pull/2161) [`f963237`](https://github.com/stella/folio/commit/f963237b6ec7e3d22ec56119ba21369202e2c5c8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Declarations are generated from explicit export types. `@stll/folio-agents` now types the paragraph alignment and line spacing rule enums in its operation schemas.
- Updated dependencies [[`f963237`](https://github.com/stella/folio/commit/f963237b6ec7e3d22ec56119ba21369202e2c5c8)]:
  - @stll/utf8-boundaries@0.1.2

## 0.39.0

### Minor Changes

- [#2100](https://github.com/stella/folio/pull/2100) [`14adc3f`](https://github.com/stella/folio/commit/14adc3f7b25721930eb926f54cc6fd2f2f868a46) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored East Asian automatic-spacing toggles through paragraph style resolution, editing, and DOCX serialization.

- [#1957](https://github.com/stella/folio/pull/1957) [`c06cd22`](https://github.com/stella/folio/commit/c06cd22e3eb924f8c2b589b5651839b9f6045246) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Draw cached line, pie, and doughnut charts with bounded path geometry, and reserve chart fallback images for explicit drawing fallbacks.

- [#1904](https://github.com/stella/folio/pull/1904) [`04cd800`](https://github.com/stella/folio/commit/04cd800a77664daedbb79b69e0015ddebdaf17c4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Section break and footnote and endnote commands run through canonical document operations in canonical sessions, including tracked note-reference revisions in suggesting mode; every inserted note paragraph receives a fresh paragraph id. The no-op `extractSelectionContext` editor command is removed: call the exported `extractSelectionContext(state)` function instead. `@stll/docx-core/ops` exports `paragraphIdsIn`.

- [#2064](https://github.com/stella/folio/pull/2064) [`4adccc0`](https://github.com/stella/folio/commit/4adccc0bdcc8d32bba87e4800c7e0ecbbbed90c8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor legacy `w:noColumnBalance` settings in continuous multi-column sections.

- [#1907](https://github.com/stella/folio/pull/1907) [`e9dc457`](https://github.com/stella/folio/commit/e9dc457d2fe49c8e124884569ffdac62c3fa06c5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose and use one pre-batch operation ordering contract for batch execution and replay.

### Patch Changes

- [#2145](https://github.com/stella/folio/pull/2145) [`3508c2f`](https://github.com/stella/folio/commit/3508c2f3d76333d5ecd558add9fd8e57ea9baa7b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving several revisions at once matches resolving them one operation at a time: a range boundary kept from a removed insertion stays at its depth beside restored inline content controls and later insertions made inside the removed content.

- [#2058](https://github.com/stella/folio/pull/2058) [`2022dbe`](https://github.com/stella/folio/commit/2022dbed10cf101a7bb9f6096a5382238e85e3c7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked move's range markers written between blocks are kept as typed blocks, in the document model and the editor, and are saved back between the same two blocks.

- [#2141](https://github.com/stella/folio/pull/2141) [`f63420f`](https://github.com/stella/folio/commit/f63420ffe0a507c789376c1ddbb1180811bc91fc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Declare paragraphs updated by revision range normalization as receiver effects.

- [#2040](https://github.com/stella/folio/pull/2040) [`e502e85`](https://github.com/stella/folio/commit/e502e855437ab1e3050ba6c0af836ee7ed4f22cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve unchanged paragraph identity when rejecting whole-document revisions.

- [#1959](https://github.com/stella/folio/pull/1959) [`8141429`](https://github.com/stella/folio/commit/814142966387495d302602c0eec00d70dee36217) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep empty comment lists and empty paragraph formatting out of the document model.

- [#2135](https://github.com/stella/folio/pull/2135) [`ade929d`](https://github.com/stella/folio/commit/ade929dda75e5f72415481feca09feccba7532fa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep comment replies in step with their root thread during editor projection and save.

- [#2030](https://github.com/stella/folio/pull/2030) [`c3f3afc`](https://github.com/stella/folio/commit/c3f3afcdef4671407853ddd011b44469856a4d68) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore explicit empty paragraph property sets when applying inverse operations.

- [#2133](https://github.com/stella/folio/pull/2133) [`2270e1a`](https://github.com/stella/folio/commit/2270e1ae47d37accbe475dcf6c6dd4003ac1fc43) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve comment endpoint ownership across inline content controls.

- [#2137](https://github.com/stella/folio/pull/2137) [`2c1d144`](https://github.com/stella/folio/commit/2c1d1448d65c53c28d26a05a03d8e8c83cc8830d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retain comment references as paragraph carriers and preserve mixed carrier order when removing a paragraph.

- [#2119](https://github.com/stella/folio/pull/2119) [`916bb0e`](https://github.com/stella/folio/commit/916bb0e84c00b505e246df165fe30058d227f07a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve copied range-boundary ownership through revision resolution and editor round trips.

- [#2012](https://github.com/stella/folio/pull/2012) [`d930da8`](https://github.com/stella/folio/commit/d930da8cae1e272a905c085bbc295abae84e1088) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Classify replacements that address only an automatic note marker as no-change edits.

- [#2060](https://github.com/stella/folio/pull/2060) [`8069ea6`](https://github.com/stella/folio/commit/8069ea6e0435ab08204ec4d140b276341e4bc68d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Update runtime dependencies.

- [#1973](https://github.com/stella/folio/pull/1973) [`f64a269`](https://github.com/stella/folio/commit/f64a26912cbf8807ae5f561a129d7d1cf667ac1b) Thanks [@dependabot](https://github.com/apps/dependabot)! - Update runtime dependencies.

- [#2129](https://github.com/stella/folio/pull/2129) [`2a6c2f8`](https://github.com/stella/folio/commit/2a6c2f888f1aea4ad20f4ba8cbae07dd19f1f307) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve directional range-boundary placement through rejected source-cut wrappers.

- [#2069](https://github.com/stella/folio/pull/2069) [`ed2a724`](https://github.com/stella/folio/commit/ed2a724d5a32cc794713c072a9bf7d9cba2816f8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep empty list-marker projections consistent with numbering definitions and legal-numbering mode.

- [#2159](https://github.com/stella/folio/pull/2159) [`2d71bf9`](https://github.com/stella/folio/commit/2d71bf9e8cd96abe15fdac82a4f75af85f415764) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Insertions at one gap run together in a batch even when a paragraph-property write ties between them, so they inherit their anchor's properties alike.

- [#2139](https://github.com/stella/folio/pull/2139) [`3b740a9`](https://github.com/stella/folio/commit/3b740a9325412fccaa925dcb8b280282cfa54693) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve excluded cached page breaks when replacing text with explicit endpoint gaps.

- [#1962](https://github.com/stella/folio/pull/1962) [`60f127c`](https://github.com/stella/folio/commit/60f127c35b47d8fd7fb4e2fa13e4c9b6ff102602) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align paragraph-mark carriers between direct deletion and accepted tracked deletion.

- [#1962](https://github.com/stella/folio/pull/1962) [`60f127c`](https://github.com/stella/folio/commit/60f127c35b47d8fd7fb4e2fa13e4c9b6ff102602) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Treat marker-only transparent inline containers as empty when resolving paragraph marks.

- [#1968](https://github.com/stella/folio/pull/1968) [`04ad8d7`](https://github.com/stella/folio/commit/04ad8d7dadea7935f7e0ccc9a9489dcd7540e5fa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep point comments anchored consistently across review reads and saves.

- [#2109](https://github.com/stella/folio/pull/2109) [`9b9396d`](https://github.com/stella/folio/commit/9b9396da89dd69d65b5ccb80d014f869981e28a1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve comment range boundaries authored inside hyperlinks.

- [#2118](https://github.com/stella/folio/pull/2118) [`65d534e`](https://github.com/stella/folio/commit/65d534ee19e1ac4dac8843cd55e48d9d4d64b463) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve move-range positions authored inside hyperlinks through editing and revision resolution.

- [#2155](https://github.com/stella/folio/pull/2155) [`d356993`](https://github.com/stella/folio/commit/d3569937d36c3ebe35065ceb1160c40c6dee6694) Thanks [@jan-kubica](https://github.com/jan-kubica)! - When rejecting leaves a kept range start after its kept end, the boundary kept from the later revision moves, in the engine and the editor alike, so the earlier one stays where it was whatever order the revisions are rejected in.

- [#1958](https://github.com/stella/folio/pull/1958) [`d2424e2`](https://github.com/stella/folio/commit/d2424e296bc4cce3cb19814048ca188de08f85d4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Derive exhausted comment ranges from their surviving coverage without changing document-operation history.

- [#2090](https://github.com/stella/folio/pull/2090) [`9174ef4`](https://github.com/stella/folio/commit/9174ef4cf41b84fb002a054ea1103a1091b651c3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An operation that removes the last reference to a note reports every paragraph record it replaces as touched, including one replaced by an equal record.

- [#1964](https://github.com/stella/folio/pull/1964) [`5f3d53a`](https://github.com/stella/folio/commit/5f3d53a72afc62bc1a78177ba159d09d78f35c87) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep current tracked style numbering authoritative and apply start overrides to full numbering-level overrides.

- [#1964](https://github.com/stella/folio/pull/1964) [`5f3d53a`](https://github.com/stella/folio/commit/5f3d53a72afc62bc1a78177ba159d09d78f35c87) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve paragraph list kind, labels, levels, and legal numbering from one canonical numbering owner across readers.

- [#2047](https://github.com/stella/folio/pull/2047) [`ba1ccf4`](https://github.com/stella/folio/commit/ba1ccf40d600babc415b11edfc7f6bafbf934e34) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The editor and the engine decide what a paragraph keeps when its mark goes with one content rule, so accepting or rejecting a paragraph mark keeps the bookmarks and comment anchors it holds in both.

- [#2112](https://github.com/stella/folio/pull/2112) [`cd077a3`](https://github.com/stella/folio/commit/cd077a386784dbd8fcdc8f04324f292332aa98d6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Derive retained-anchor receiver effects from the canonical paragraph-boundary removal plan.

- [#1958](https://github.com/stella/folio/pull/1958) [`d2424e2`](https://github.com/stella/folio/commit/d2424e296bc4cce3cb19814048ca188de08f85d4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve automatic note markers within note-story text replacements.

- [#2107](https://github.com/stella/folio/pull/2107) [`10b1a2d`](https://github.com/stella/folio/commit/10b1a2da617ba59cd0af8ace022bc288b87617a4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share linear range multiplicity classification and preserve source ownership when rejecting revisions around range boundaries.

- [#2057](https://github.com/stella/folio/pull/2057) [`c9fa99c`](https://github.com/stella/folio/commit/c9fa99c57eaa34010c4076ef744ff2bdd828c07a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rejecting a tracked change keeps the bookmark, comment and move range boundaries its content held where that content was, in the engine and in the editor, and keeps every such range starting before it ends.

- [#2139](https://github.com/stella/folio/pull/2139) [`3b740a9`](https://github.com/stella/folio/commit/3b740a9325412fccaa925dcb8b280282cfa54693) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve replacement source-run joins around cached page breaks without merging separately authored runs.

- [#1958](https://github.com/stella/folio/pull/1958) [`d2424e2`](https://github.com/stella/folio/commit/d2424e296bc4cce3cb19814048ca188de08f85d4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove range comments when edits leave none of their covered content, while preserving authored points and ranges over inline objects.

- [#2132](https://github.com/stella/folio/pull/2132) [`ae3ff3d`](https://github.com/stella/folio/commit/ae3ff3d84271df27b9b17f41b9b1156fd2092693) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Declare canonical paragraph receivers while planning revision resolution, including transitive table joins.

- [#1962](https://github.com/stella/folio/pull/1962) [`60f127c`](https://github.com/stella/folio/commit/60f127c35b47d8fd7fb4e2fa13e4c9b6ff102602) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Unify paragraph-mark refusal reasons before section-ending tables.

- [#2124](https://github.com/stella/folio/pull/2124) [`7830253`](https://github.com/stella/folio/commit/7830253614227c6fcda0a86a06bc5399249ec798) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse insertions that create inverted bookmark or move ranges.

- [#2110](https://github.com/stella/folio/pull/2110) [`a75c101`](https://github.com/stella/folio/commit/a75c1010e549b8fe171b2235b3e04165c1fddca7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep retained range boundaries at their source seam when batch rejection removes both cut containers.

- [#1964](https://github.com/stella/folio/pull/1964) [`5f3d53a`](https://github.com/stella/folio/commit/5f3d53a72afc62bc1a78177ba159d09d78f35c87) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Inserted and merged paragraphs keep the list they ask for: a direct merge into a paragraph pending deletion, a tracked insertion before deleted closing paragraphs, and an insertion after a table in a batch that empties the closing paragraph now read as applying the operations one at a time or directly does. A merge that sets the merged paragraph's properties conflicts with another property edit of that paragraph in the same batch, a suggested operation whose id names a still pending suggestion gets a suggestion id of its own, and clearing a numbered paragraph style drops the style's numbering from the document model. A level-only override with no style numbering is no list in any reader, and clean Markdown reads a pending numbering removal as accepted.

- [#2034](https://github.com/stella/folio/pull/2034) [`4b26e14`](https://github.com/stella/folio/commit/4b26e145ef7351451caa45b9641529f945addac1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep section properties when inserted paragraph boundaries are rejected.

- [#2033](https://github.com/stella/folio/pull/2033) [`6713246`](https://github.com/stella/folio/commit/6713246807a38ccce24a872683dfdb0a60f74a2c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep comment boundary ownership through revision resolution and whole-range deletion.

- [#1968](https://github.com/stella/folio/pull/1968) [`04ad8d7`](https://github.com/stella/folio/commit/04ad8d7dadea7935f7e0ccc9a9489dcd7540e5fa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep both comment range boundaries and references at their authored tracked-change positions.

- [#2092](https://github.com/stella/folio/pull/2092) [`c34806f`](https://github.com/stella/folio/commit/c34806fff64bc338c5e7917bbfe8d69e094b67f2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Synchronize text-only edits in text-box alternate-content fallbacks.

- [#2130](https://github.com/stella/folio/pull/2130) [`723f1c4`](https://github.com/stella/folio/commit/723f1c4a984bc6a25bb53881bb0c78f7ca967441) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve comment anchors when a removed paragraph can only join the following table.
- Updated dependencies [[`20ee7f4`](https://github.com/stella/folio/commit/20ee7f4e19b0d4e3f47a8d72351ea9d288a5dcb2)]:
  - @stll/utf8-boundaries@0.1.1

## 0.38.0

### Minor Changes

- [#1683](https://github.com/stella/folio/pull/1683) [`ab765b7`](https://github.com/stella/folio/commit/ab765b79049f388c63b02ff630521b140b01d91b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve aligned comment coverage, retained point comments, and inherited table-cell formatting across editing and revision resolution.

  Expose in-memory document resolution snapshots and keep section content aligned with the edited body.

  Keep comment relationships aligned with modeled comment threads.

### Patch Changes

- [#1705](https://github.com/stella/folio/pull/1705) [`7fbd519`](https://github.com/stella/folio/commit/7fbd519345dd2e3ca3c09eeb0c116e059064aecb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Saving could corrupt astral characters such as emoji at 16 KiB boundaries inside document XML parts. Upgrade to keep them intact.

- [#1696](https://github.com/stella/folio/pull/1696) [`82ea0f0`](https://github.com/stella/folio/commit/82ea0f074f59d5e735d02d207462e18dac97fba3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose the published ESM module targets to synchronous module loaders.
- Updated dependencies [[`7fbd519`](https://github.com/stella/folio/commit/7fbd519345dd2e3ca3c09eeb0c116e059064aecb)]:
  - @stll/utf8-boundaries@0.1.0

## 0.37.0

### Minor Changes

- [#1684](https://github.com/stella/folio/pull/1684) [`19a80db`](https://github.com/stella/folio/commit/19a80db2fd8b3bc2e32f04586c12c63a36521d59) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve input boundaries and surviving paragraph properties, preserve prose formatting toggles, and coalesce pending run formatting revisions.

### Patch Changes

- [#1703](https://github.com/stella/folio/pull/1703) [`e03e0ec`](https://github.com/stella/folio/commit/e03e0ec4ff4185a8f15f7dc5aad513f7e6b72588) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep retained revision identities aligned when pending run formatting is replaced.

- [#1697](https://github.com/stella/folio/pull/1697) [`73ec846`](https://github.com/stella/folio/commit/73ec846b9193c5d522b1d263f8813ed36fd099fa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open documents with repeated bookmark IDs and report imported identity collisions.

## 0.36.3

### Patch Changes

- [#1661](https://github.com/stella/folio/pull/1661) [`93bb3b8`](https://github.com/stella/folio/commit/93bb3b8a1b85c014c0e02b6d40b5f1d8384e53ce) Thanks [@dependabot](https://github.com/apps/dependabot)! - Regenerate the document kernel and text shaper bindings for updated Rust dependencies.

- [#1662](https://github.com/stella/folio/pull/1662) [`2c966a4`](https://github.com/stella/folio/commit/2c966a474bf865013a5342ef0460b65f766ab2af) Thanks [@dependabot](https://github.com/apps/dependabot)! - Use the UTF-8 Quick-XML API while preserving namespace, text normalization, and Flat OPC projection behavior.

## 0.36.2

### Patch Changes

- [#1658](https://github.com/stella/folio/pull/1658) [`d38384d`](https://github.com/stella/folio/commit/d38384d6d8654fa5c4a95fc300683c993bcd7771) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Regenerate bundled WebAssembly with the pinned Rust toolchain.

## 0.36.1

### Patch Changes

- [#1605](https://github.com/stella/folio/pull/1605) [`8919807`](https://github.com/stella/folio/commit/891980702cec4acd0ea26f19534eb78da4741f43) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize canonical list rendering after every edit, record authored paragraph review snapshots, and preserve section views when replacing body content.

## 0.36.0

### Minor Changes

- [#1642](https://github.com/stella/folio/pull/1642) [`3a53991`](https://github.com/stella/folio/commit/3a53991f2229737a217591f23e81ea7a7004d342) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Breaking: projection wire schema 7 adds the paragraph container (body or tableCell), retained in fragments without document coordinates, and alignment completeness. Document-default alignment uses distinct docDefaults provenance; unread dependencies and unsupported alignment values remain explicit unknowns.

## 0.35.0

### Minor Changes

- [#1637](https://github.com/stella/folio/pull/1637) [`68038ab`](https://github.com/stella/folio/commit/68038abbcae49f963de34ebbbf6b1897ef69d6e5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Replace the global projection formatting status with per-family completeness in wire schema 6. Read the status for each formatting family; direct highlight remains known when styles are unavailable.

  Expose the initialized Rust runtime's projection schema version through `docxProjectionSchemaVersion`.

## 0.34.0

### Minor Changes

- [#1628](https://github.com/stella/folio/pull/1628) [`1e24558`](https://github.com/stella/folio/commit/1e24558cf0936153a89023da654072a0e5428768) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose bounded main-document XML and relationship-selected Flat OPC projection through the shared kernel and WebAssembly wire.

- [#1631](https://github.com/stella/folio/pull/1631) [`4990c8b`](https://github.com/stella/folio/commit/4990c8bb2f4ba6fff3ecfdab179b4d64045481f4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose bounded single-paragraph Flat OPC projection through the shared kernel and WebAssembly wire. Document structural evidence is explicitly unknown and fragment-local table coordinates are unavailable.

### Patch Changes

- [#1626](https://github.com/stella/folio/pull/1626) [`3c06b3d`](https://github.com/stella/folio/commit/3c06b3d22a152478a43c60715e7664ffbea87d10) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project semantic highlight spans only from direct run markup, excluding paragraph and character style inheritance while preserving other resolved formatting.

## 0.33.0

### Minor Changes

- [#1538](https://github.com/stella/folio/pull/1538) [`2ce4891`](https://github.com/stella/folio/commit/2ce489111e78badb8934d95216679c7da7b03ca4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route canonical comment lifecycles through semantic document operations with exact journal inverses and package-owned anchors.

### Patch Changes

- [#1538](https://github.com/stella/folio/pull/1538) [`2ce4891`](https://github.com/stella/folio/commit/2ce489111e78badb8934d95216679c7da7b03ca4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accept replies to loaded replies without own body markers, and map comment boundaries carried by marks during canonical editing.

- [#1541](https://github.com/stella/folio/pull/1541) [`ce048e0`](https://github.com/stella/folio/commit/ce048e05bf6b0684462dcf1c55f1382fc035a2f9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep zero-width markers and empty runs out of canonical input formatting affinity.

- [#1538](https://github.com/stella/folio/pull/1538) [`2ce4891`](https://github.com/stella/folio/commit/2ce489111e78badb8934d95216679c7da7b03ca4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Anchor revision-associated comments on empty paragraphs and preserve exact undo and redo.

- [#1538](https://github.com/stella/folio/pull/1538) [`2ce4891`](https://github.com/stella/folio/commit/2ce489111e78badb8934d95216679c7da7b03ca4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicitly undefined owned story maps and section properties through lifecycle undo and redo.

- [#1606](https://github.com/stella/folio/pull/1606) [`6b7f7a6`](https://github.com/stella/folio/commit/6b7f7a67d394f36189cd68028ede02d9655a5a54) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve tracked child revisions in simple fields and serialize tracked hyperlink field results.

- [#1538](https://github.com/stella/folio/pull/1538) [`2ce4891`](https://github.com/stella/folio/commit/2ce489111e78badb8934d95216679c7da7b03ca4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open documents whose root comments have no source anchor in canonical sessions, keeping those comments read-only and preserved on save.

## 0.32.0

### Minor Changes

- [#1472](https://github.com/stella/folio/pull/1472) [`2f3acb1`](https://github.com/stella/folio/commit/2f3acb1799c26bc943f84214888d0ce85133c55e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compile clipboard paste and moves into atomic canonical edits with package resource imports and exact history.

  Clipboard package imports use schema 9 over the shipped schema-8 table contract; older operation and batch envelopes receive structured refusals.

- [#1582](https://github.com/stella/folio/pull/1582) [`c5a0926`](https://github.com/stella/folio/commit/c5a0926238207aaf2d0b98caaf162b15dda0589c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore prior direct properties in Original projections and expose located table property exception revisions.

- [#1568](https://github.com/stella/folio/pull/1568) [`baeff8a`](https://github.com/stella/folio/commit/baeff8aa3dc1a01a0bd13d6d92fa857653f21549) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Locate tracked paragraph-mark revisions with a zero-width span and the new `"paragraph-mark"` revision content kind, and keep revision and comment locations known when the selected revision view joins paragraphs.

### Patch Changes

- [#1472](https://github.com/stella/folio/pull/1472) [`2f3acb1`](https://github.com/stella/folio/commit/2f3acb1799c26bc943f84214888d0ce85133c55e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve copied paragraph review and surviving mark formatting, share edited run seams between clipboard modes, and apply canonical clipboard transforms and composition lifecycle.

- [#1537](https://github.com/stella/folio/pull/1537) [`cea1b5f`](https://github.com/stella/folio/commit/cea1b5ff41ae983617a1f2a42444eadcf41d617e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Lower hyperlink commands to canonical editor intents with exact inverses and explicit tracked-mode refusals.

- [#1589](https://github.com/stella/folio/pull/1589) [`00b65d4`](https://github.com/stella/folio/commit/00b65d49c0bc0867bd5d3557b67f2d8c1f14aab3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve source marker ordinals when activating canonical inline projections.

- [#1592](https://github.com/stella/folio/pull/1592) [`836731a`](https://github.com/stella/folio/commit/836731aed2040950a49a12abd7f9049aeb0479b6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve source range anchors in planned replacements, transfer them once in editing moves, and retain them at source during suggested moves.

- [#1539](https://github.com/stella/folio/pull/1539) [`38340e2`](https://github.com/stella/folio/commit/38340e2e66203b84010fed1b2a2d7092eedec2d6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate tables of contents through canonical operations while preserving heading bookmark references and exact history.

- [#1472](https://github.com/stella/folio/pull/1472) [`2f3acb1`](https://github.com/stella/folio/commit/2f3acb1799c26bc943f84214888d0ce85133c55e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep source container joins explicit for tracked replacement plans and type clipboard cut events.

- [#1472](https://github.com/stella/folio/pull/1472) [`2f3acb1`](https://github.com/stella/folio/commit/2f3acb1799c26bc943f84214888d0ce85133c55e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve clipboard identity, relationship and story projection ownership.

- [#1472](https://github.com/stella/folio/pull/1472) [`2f3acb1`](https://github.com/stella/folio/commit/2f3acb1799c26bc943f84214888d0ce85133c55e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep header and footer source captures across document clones, so an edited part saved from a clone reports its source mismatch.

- [#1581](https://github.com/stella/folio/pull/1581) [`f2ddacf`](https://github.com/stella/folio/commit/f2ddacfe536f9f1f9661c7320e3eaee89f53b176) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve effective table and paragraph style cascades, preserve default paragraph styles for invalid references, and keep numbering-label formatting separate from paragraph content.

- [#1547](https://github.com/stella/folio/pull/1547) [`310afcd`](https://github.com/stella/folio/commit/310afcdda4b2b40c2c2d4d93697c693ca66aa53d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore captured section story snapshots exactly when they differ from their referenced package parts.

- [#1573](https://github.com/stella/folio/pull/1573) [`ef1117d`](https://github.com/stella/folio/commit/ef1117d877a9040b6ae88ea25df0886204feda0b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Join paragraphs across tracked moved paragraph marks: a `w:moveFrom` mark removes the paragraph break in the current view like a `w:del` mark, and a `w:moveTo` mark removes it in the original view like a `w:ins` mark. Materialize `w:endnoteReference` as U+0002 in projected text, like `w:footnoteReference`.

- [#1575](https://github.com/stella/folio/pull/1575) [`52b6075`](https://github.com/stella/folio/commit/52b6075528a05afe3115a315955f5b384f3d570b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep review and structural facts known across paragraph joins and hidden content, and coalesce joined formatting spans.

- [#1579](https://github.com/stella/folio/pull/1579) [`8a7231b`](https://github.com/stella/folio/commit/8a7231bcb1fd45b8ced5dbe9f407a9a037db6f0b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Locate table, grid, row, and cell property revisions over their owning content.

- [#1588](https://github.com/stella/folio/pull/1588) [`a74f206`](https://github.com/stella/folio/commit/a74f2065839ad396943fc66820e454083f1522d2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write DOCX package parts without creating parent directory entries, using deterministic default ZIP dates.

- [#1555](https://github.com/stella/folio/pull/1555) [`3e389e3`](https://github.com/stella/folio/commit/3e389e32d3cd992c70f1dfe6f025498bc3c0b3cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve content-control source cuts when rejecting chains of inserted paragraph breaks.

## 0.31.0

### Minor Changes

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add shared structural and formatting intent compilation with reversible numbering instance and section endpoint operations.

- [#1331](https://github.com/stella/folio/pull/1331) [`2f9695f`](https://github.com/stella/folio/commit/2f9695f96f282239d6cc6b46791606c1faa15bff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add stable header, footer, and note story addresses, story lifecycle and section-property operations, and exact inverses in document operation schema version 5.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compile editor text and paragraph intents to direct or tracked operations, preserving cross-container text ranges, existing reviews, and section metadata through exact inverses in operation schema 5.

- [#1474](https://github.com/stella/folio/pull/1474) [`e953583`](https://github.com/stella/folio/commit/e95358379352f16aa2196d5ec95b448b02bfe692) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add semantic table operations, tracked table resolution and exact structural inverses.

  Table operations require exclusive edits and are refused in sequenced batches. Schema 8 envelopes preserve the complete operation contract; older schemas receive structured refusals.

- [#1326](https://github.com/stella/folio/pull/1326) [`e6b890c`](https://github.com/stella/folio/commit/e6b890cd058463cb0efdf66badd94c1a6c97ca88) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add versioned operation batches, conservative rebasing, and deterministic in-memory sequencing and optimistic clients.

### Patch Changes

- [#1331](https://github.com/stella/folio/pull/1331) [`2f9695f`](https://github.com/stella/folio/commit/2f9695f96f282239d6cc6b46791606c1faa15bff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound lifecycle delta inference while preserving exact story inverses.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Omit absent editing allocation fields when compiling canonical document operations.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve review identity provenance when cutting retained deletions and omit unused paragraph formatting allocations.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve exact section and authored run history, capture undefined fields losslessly in schema-7 operation journals, and return atomic refusals for unsupported envelopes.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share plain-run seam merging between parsing and canonical joins, restore exact undo content, and preserve optional identity pools.

- [#1331](https://github.com/stella/folio/pull/1331) [`2f9695f`](https://github.com/stella/folio/commit/2f9695f96f282239d6cc6b46791606c1faa15bff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Correct shared story edit refusals, lifecycle inverses, and package serialization.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored run-property presence and journal parser-compatible text seams across canonical story edits.

- [#1335](https://github.com/stella/folio/pull/1335) [`3ac5183`](https://github.com/stella/folio/commit/3ac5183e0e51648dd8392584269e3498170c5bda) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Load generated schema facts without expanding every data value into compiler types.

- [#1474](https://github.com/stella/folio/pull/1474) [`e953583`](https://github.com/stella/folio/commit/e95358379352f16aa2196d5ec95b448b02bfe692) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retain consumed identities in compiled tracked plans and prepare canonical table intents.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve absent run formatting and classify every canonical input and package operation.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use the same visible paragraph group to select formatting for direct and tracked editor joins.

- [#1335](https://github.com/stella/folio/pull/1335) [`3ac5183`](https://github.com/stella/folio/commit/3ac5183e0e51648dd8392584269e3498170c5bda) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve OPC default content types from the final filename segment and require overrides for extensionless parts.

- [#1451](https://github.com/stella/folio/pull/1451) [`4c286c5`](https://github.com/stella/folio/commit/4c286c5b7c898d4436b8968f20fa8f0beea0e976) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve folded inline LISTNUM display text when resolving list markers.

- [#1335](https://github.com/stella/folio/pull/1335) [`3ac5183`](https://github.com/stella/folio/commit/3ac5183e0e51648dd8392584269e3498170c5bda) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Validate saved OOXML attribute values, content types, and package relationships.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound editor identity allocation to the intent and retain only consumed IDs in operation history.

- [#1470](https://github.com/stella/folio/pull/1470) [`7c35edd`](https://github.com/stella/folio/commit/7c35eddd79bdf2c08805e4468e58daf42fb178f7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the list-number fields shown in a paragraph's list marker when the paragraph is saved. Such a field, and the tab after it, now stay in the paragraph content as preserved markup and are written back as they were read.

  The reader now draws a list-number field in the list marker only when the field opens a numbered paragraph whose marker has text of its own. A list-number field that follows other text, or that stands in a bulleted paragraph, is shown inline where it stands.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve one paragraph property revision across tracked edits and keep unused editor operation exports internal.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep existing folded paragraph-review ids out of newly recorded revision results.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use shared numbering readers for editor intent validation and list-level resolution.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored empty paragraph property sets through tracked structural editing.

- [#1474](https://github.com/stella/folio/pull/1474) [`e953583`](https://github.com/stella/folio/commit/e95358379352f16aa2196d5ec95b448b02bfe692) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored inline boundaries when accepting replacements that empty an older review wrapper.

- [#1511](https://github.com/stella/folio/pull/1511) [`d52874d`](https://github.com/stella/folio/commit/d52874da2a0da3836d090032cffeae1473e7ba99) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retain private paragraph formatting captures across model edits and their inverses.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align batch schema fixtures and explicitly refuse unsupported paragraph-review and section-boundary wire metadata.

- [#1331](https://github.com/stella/folio/pull/1331) [`2f9695f`](https://github.com/stella/folio/commit/2f9695f96f282239d6cc6b46791606c1faa15bff) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align batch schema fixtures with document operations and explicitly refuse unsupported story lifecycle sequencing pairs.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore recorded source run seams when rejecting replacements that inserted paragraph breaks.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve canonical review resolution and explicit formatting presence across editor and document boundaries.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve tracked source seams across review order and complete run review policy handling.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve section endpoints case-insensitively and return typed refusals for non-object journal operations.

- [#1326](https://github.com/stella/folio/pull/1326) [`e6b890c`](https://github.com/stella/folio/commit/e6b890cd058463cb0efdf66badd94c1a6c97ca88) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share the no-effect discriminator between sequencing capture and wire validation.

- [#1326](https://github.com/stella/folio/pull/1326) [`e6b890c`](https://github.com/stella/folio/commit/e6b890cd058463cb0efdf66badd94c1a6c97ca88) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reject cross-paragraph ranges before rebasing and report client rejections once.

- [#1446](https://github.com/stella/folio/pull/1446) [`9b5a8ff`](https://github.com/stella/folio/commit/9b5a8ffe7bbbdb06a1e24231c33280b473cf3356) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve pending source fragments when resolving split paragraphs around later insertions.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve canonical suggestion content across save and reopen, minimize operation exports, and reduce type instantiation cost.

- [#1474](https://github.com/stella/folio/pull/1474) [`e953583`](https://github.com/stella/folio/commit/e95358379352f16aa2196d5ec95b448b02bfe692) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Budget final-column deletion identities from table content and trim unused table intent allocations.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Report stable folded paragraph review ids and allow terminal deletion to retain the original review baseline.

- [#1339](https://github.com/stella/folio/pull/1339) [`2fafe42`](https://github.com/stella/folio/commit/2fafe426a9abaa9fd19ed0c01ca39c5431db2c5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Merge plain runs at accepted paragraph joins consistently with direct joins and refuse unrepresentable tracked merge depths.

- [#1335](https://github.com/stella/folio/pull/1335) [`3ac5183`](https://github.com/stella/folio/commit/3ac5183e0e51648dd8392584269e3498170c5bda) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Correct XML entity decoder augmentation imports for strict typechecking.

## 0.30.1

### Patch Changes

- [#1333](https://github.com/stella/folio/pull/1333) [`c3a3a38`](https://github.com/stella/folio/commit/c3a3a3802ca05edc1898394b0b49e08e9d59185a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allocate bounded OOXML identifiers across editors and reviewers, normalize legacy imported identities consistently, and reject out-of-range identifiers when saving DOCX files.

- [#1333](https://github.com/stella/folio/pull/1333) [`c3a3a38`](https://github.com/stella/folio/commit/c3a3a3802ca05edc1898394b0b49e08e9d59185a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Skip namespace collection and reservation for in-range imports; reuse repaired document trees, scan other imported identifiers without retaining XML trees, and compress repaired baselines at the fastest DEFLATE level.

## 0.30.0

### Minor Changes

- [#1283](https://github.com/stella/folio/pull/1283) [`a91219c`](https://github.com/stella/folio/commit/a91219c5d82e1ed5ce8dcd68695b98326e7b51b9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add reversible direct and tracked paragraph deletion with container-final identity preservation.

### Patch Changes

- [#1292](https://github.com/stella/folio/pull/1292) [`cba570e`](https://github.com/stella/folio/commit/cba570e40ecea26e79fc9c81ced1803c549a9793) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Accepting tracked changes one operation at a time now gives the same result as accepting them at once: a record an acceptance empties goes, as a direct deletion leaves none, instead of staying empty beside the piece of it a later acceptance brings back. Rejecting a paste that cut a content control, hyperlink or other inline container now gives the container back whole: the piece the rejection empties folds into the other piece even when no resolved change lies between them.

- [#1291](https://github.com/stella/folio/pull/1291) [`b56dcb9`](https://github.com/stella/folio/commit/b56dcb95631430e52f2fba8103e9d824e277e481) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Share comment-anchor classification across tracked deletion planning and wrapping.

- [#1281](https://github.com/stella/folio/pull/1281) [`ad0dca4`](https://github.com/stella/folio/commit/ad0dca442e5fb200f60dea720aeae060996e1b70) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add direct table insertion and deletion operations with exact structural inverses.

- [#1289](https://github.com/stella/folio/pull/1289) [`0781905`](https://github.com/stella/folio/commit/0781905244c23e2192b38a8d7cb2d09c486b2fc6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Address row inverses through the containing table when a surviving cell starts with a nested table.

- [#1288](https://github.com/stella/folio/pull/1288) [`b18d455`](https://github.com/stella/folio/commit/b18d45571cb08c29936d9942dcee804bdf09b37a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Plan tracked deletion and replacement across adjacent paragraphs with shared revision allocation.

- [#1299](https://github.com/stella/folio/pull/1299) [`0c2c665`](https://github.com/stella/folio/commit/0c2c6654e5d943ba4618beba98a8bc9f77b8c00a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked split whose new paragraph is the second half now refuses mark run properties other than the paragraph's own. That half ends with the paragraph's own mark, and no property change records its run properties, so rejecting the split could not give them back.

- [#1287](https://github.com/stella/folio/pull/1287) [`8c1ce41`](https://github.com/stella/folio/commit/8c1ce41608f848f1e565da5ae16ebf85f5242821) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Support tracked whole-table insertion and deletion with cell paragraph marks and terminal table insertion carriers.

## 0.29.0

### Minor Changes

- [#1212](https://github.com/stella/folio/pull/1212) [`f545ef1`](https://github.com/stella/folio/commit/f545ef1526264870550ae4300288d43b5c48d5cd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `resolveRevision` to `@stll/docx-core/ops`: accept or reject tracked insertions, deletions, moves, run and paragraph property changes and paragraph marks by revision id, applied as primitive operations with an exact inverse.

- [#1211](https://github.com/stella/folio/pull/1211) [`3a7d1c0`](https://github.com/stella/folio/commit/3a7d1c0ee2f0b9b9cbfbf625a9daa8aea00c7592) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add tracked changes to `@stll/docx-core/ops`: a `revision` stamp on the text, formatting and paragraph operations records the edit as a tracked insertion, deletion, property change or paragraph mark, with an exact inverse. `planTrackedDeletion` plans a tracked deletion around the author's own insertions, and `revisionIdDemand` counts the revision ids an operation takes. The operation schema is now version 2.

- [#1257](https://github.com/stella/folio/pull/1257) [`c80cb40`](https://github.com/stella/folio/commit/c80cb40f19165f466da9bfb113798ba2446167bc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add direct and tracked paragraph insertion with reversible review and operation schema version 3.

- [#1263](https://github.com/stella/folio/pull/1263) [`5c625cb`](https://github.com/stella/folio/commit/5c625cbfd950906c1e0cd6d357fb3af6def0baee) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add direct and tracked table-row operations with exact undo and revision resolution.

### Patch Changes

- [#1279](https://github.com/stella/folio/pull/1279) [`d5a9ba3`](https://github.com/stella/folio/commit/d5a9ba3d0a9500afaf76ad720b016f43a694c10c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align shared dependency versions: jszip 3.10.2, valibot 1.5.0, @base-ui/react 1.8.0.

- [#1285](https://github.com/stella/folio/pull/1285) [`d890ce9`](https://github.com/stella/folio/commit/d890ce998ece231903b94b138d22bd3624f94f45) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound archive inflation when opening documents. A document now fails to open when one of its parts inflates past the size the archive declares for it, or when a markup part or the package as a whole exceeds the compression-ratio limit; `maxCompressionRatio` in the archive and unzip options adjusts that limit.

- [#1277](https://github.com/stella/folio/pull/1277) [`dc334e2`](https://github.com/stella/folio/commit/dc334e2256f129f1d202124bde894334254c33c3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep empty content controls when resolving tracked changes: only a container the resolution itself emptied merges into an alike neighbour, inline and across a removed paragraph mark.

- [#1244](https://github.com/stella/folio/pull/1244) [`003d387`](https://github.com/stella/folio/commit/003d387c4cabf72f160367ca536ed0757f9ae5a7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize empty revision wrappers before applying document operations, preserving meaningful zero-width markup.

- [#1271](https://github.com/stella/folio/pull/1271) [`1c68631`](https://github.com/stella/folio/commit/1c68631782a0590d53de969184e596d249fe4ba7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolving a revision now merges the seam inside the alike records it merges, so an accepted tracked insertion agrees with the direct one when an empty wrapper meets its alike neighbour inside a hyperlink.

## 0.28.1

### Patch Changes

- [#1120](https://github.com/stella/folio/pull/1120) [`225a53a`](https://github.com/stella/folio/commit/225a53ab432f1a5bbbb2ab31087b3683403c9c09) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose block-level custom XML content to readers, editing, and comparison while preserving its wrapper on save.

- [#1125](https://github.com/stella/folio/pull/1125) [`b24b60a`](https://github.com/stella/folio/commit/b24b60aa87d65f7d29808df6df210c349dc24e8b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `compileMarkdownToContent` no longer drops a table, code block, or blockquote nested inside a list item; it keeps the content as a following block and reports the flattening on the new `MarkdownContent.warnings` field.

- [#1137](https://github.com/stella/folio/pull/1137) [`1caa4ac`](https://github.com/stella/folio/commit/1caa4ac16c99b60bc223d59902eaef025ec7d635) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose content-bearing OOXML carriers to readers or report unsupported content explicitly during comparison.

- [#1173](https://github.com/stella/folio/pull/1173) [`4170c13`](https://github.com/stella/folio/commit/4170c13372d5d4a7f541f1ef727ad1b88f7585ef) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve legacy form field data through document and editor round trips.

## 0.28.0

### Minor Changes

- [#1054](https://github.com/stella/folio/pull/1054) [`d2fa7cd`](https://github.com/stella/folio/commit/d2fa7cd623b17238d89116aed3fd909549a6c272) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `@stll/docx-core/ops`: text and formatting operations on the document model (insert text, delete a range, set run or paragraph properties, split and join paragraphs), addressed by paragraph id and logical offset, each applied deterministically with an exact recorded inverse.

## 0.27.0

### Minor Changes

- [#1027](https://github.com/stella/folio/pull/1027) [`9cdabe4`](https://github.com/stella/folio/commit/9cdabe49b5dbcca8eda974042d27598df7fa8515) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Lay out and paint the text boxes inside a DrawingML group (`wpg:wgp`, nested `wpg:grpSp`) as text boxes: each child's frame is mapped through the group's `a:chOff`/`a:chExt` child coordinate space, keeps its `wps:bodyPr` insets, anchoring and autofit, and carries its own rotation and flips. Editing that text writes it back into the group on save, which keeps the group intact. The group preview also draws zero-extent lines, scales line widths into the child space and renders nested groups. Adds `Shape.groupChild` / `TextBox.groupChild` (`DrawingGroupChild`) to the model.

### Patch Changes

- [#1025](https://github.com/stella/folio/pull/1025) [`e32ad9d`](https://github.com/stella/folio/commit/e32ad9d9cde078e97481af41cdd1496022df1bd0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop a legacy `FORMCHECKBOX` field's synthesized display glyph from being written into `w:sdtContent`/the field result on save when the field carries no cached result of its own. The parser still models the glyph so the editor can paint it, flagged as a display-only fallback (`ComplexField.fieldResultIsFallback`) the serializer now honours by leaving a resultless field's result empty, matching the source. `w:sdt`/`w14:checkbox` content controls already round-tripped their authored `w:sdtContent` correctly and are unaffected.

- [#1029](https://github.com/stella/folio/pull/1029) [`2ad7e08`](https://github.com/stella/folio/commit/2ad7e085f03318182b27b07b75f29cf607475e3d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the `mc:AlternateContent` around a DrawingML shape or text box on save, `mc:Fallback` included, while the shape is unedited, through the model and the editor. An edited shape is regenerated from the model without its Fallback, which would otherwise contradict the new Choice for consumers that read only the Fallback.

- [#1026](https://github.com/stella/folio/pull/1026) [`122fdd2`](https://github.com/stella/folio/commit/122fdd29a99c1c3ae3004ad14db3f9d31e1ff1de) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Paint a text box's linear `a:gradFill`, honouring `a:lin@scaled`, and keep the gradient when the editor saves the box. Read `a:lin@scaled` into the gradient model.

## 0.26.0

### Minor Changes

- [#1021](https://github.com/stella/folio/pull/1021) [`44efb6d`](https://github.com/stella/folio/commit/44efb6dbf73f55cb22754e44bb5b3eff3ebf4e8c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Lay out and paint endnotes: they paginate with the body after its last block (`w:pos="docEnd"`) or at the end of each section (`sectEnd`, honouring `w:noEndnote`), open with a separator rule and continue under a continuation separator, show the same number as their body reference, and open the note editor on double-click. Read `w:settings/w:endnotePr`.

### Patch Changes

- [#1008](https://github.com/stella/folio/pull/1008) [`a96d9c8`](https://github.com/stella/folio/commit/a96d9c878509256ee662e1a8e5629b35dc66a62c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Send the tab after a list number to a custom tab stop between the number and `w:ind@left` instead of the hanging indent, apply the numbering level's `w:pPr/w:tabs` to its paragraphs, and honour `w:doNotUseIndentAsNumberingTabStop`.

- [#1005](https://github.com/stella/folio/pull/1005) [`d8e4211`](https://github.com/stella/folio/commit/d8e4211883b59e1b179cc0e26bcf3987eacebe00) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Format list numbers from the paragraph mark's run properties (including a `w:rStyle` character style) under the numbering level's `w:rPr`, instead of from the first text run, and paint their colour. Size an empty paragraph by a character style its mark names.

## 0.25.3

### Patch Changes

- [#962](https://github.com/stella/folio/pull/962) [`e2adc47`](https://github.com/stella/folio/commit/e2adc4749c6f8d837149324821713a04bed7ab2f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve block content controls nested directly inside table cells.

- [#957](https://github.com/stella/folio/pull/957) [`2e627aa`](https://github.com/stella/folio/commit/2e627aa74b7fa8b5ac019d4322e4f188700b2c19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep move-range boundaries inside their tracked-change wrapper across DOCX parse, editor projection, and save.

- [#966](https://github.com/stella/folio/pull/966) [`88c21aa`](https://github.com/stella/folio/commit/88c21aa9e7ac2ee1314d0023d03395b4e7c04316) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep tracked-move range boundaries inside inline content controls through editor saves.

- [#956](https://github.com/stella/folio/pull/956) [`3786fb0`](https://github.com/stella/folio/commit/3786fb015a26617434a07571015f7ad32df89fe3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve document-level page backgrounds through editor saves.

- [#960](https://github.com/stella/folio/pull/960) [`7242596`](https://github.com/stella/folio/commit/7242596401b48cc4866de941e4752f4a87668ad1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reject body note references that use the separator-record ids reserved by OOXML.

## 0.25.2

### Patch Changes

- [#948](https://github.com/stella/folio/pull/948) [`398d94d`](https://github.com/stella/folio/commit/398d94d553ed31da61a5985d576c3d264011e995) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve tracked numbering insertions when edited paragraphs are saved.

## 0.25.1

### Patch Changes

- [#943](https://github.com/stella/folio/pull/943) [`371ff8f`](https://github.com/stella/folio/commit/371ff8f13f0c04381428c2093529376affcc2cd8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored hyperlink and note-reference metadata across editor saves.

- [#944](https://github.com/stella/folio/pull/944) [`55229a3`](https://github.com/stella/folio/commit/55229a38e4906fb471749d3f5420d7d0fcc9a8d0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored numbering-format metadata for lists and note numbering.

- [#945](https://github.com/stella/folio/pull/945) [`dade6b3`](https://github.com/stella/folio/commit/dade6b363053471cb26852b1b49fc93e5e09360b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored table-cell identifiers through editor saves.

- [#949](https://github.com/stella/folio/pull/949) [`f7081b0`](https://github.com/stella/folio/commit/f7081b0bfecd5b94a2bd7df63933bfe4c4ff9b01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored carriage-return elements through editor saves.

## 0.25.0

### Minor Changes

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a bookmark boundary inside the inline content control that held it.

  `CT_SdtContentRun` reaches `w:bookmarkStart` and `w:bookmarkEnd` through `EG_RunLevelElts > EG_RangeMarkupElements`, so a marker inside `w:sdtContent` is markup Word writes. folio lifted it out to a sibling of the control. That is not a re-spelling: a bookmark whose extent was the control's content came back starting before the control, so a `REF` field or a link to it resolved to a different range, and a marker in the middle of the content split one control into two carrying the same `w:id`, `w:tag` and data binding.

  `InlineSdt["content"]` gains `BookmarkStart` and `BookmarkEnd`, and `INLINE_SDT_CONTENT` admits them: the admission map is bound to the content type, so the parser, the serializer and the editor's save filter all follow from the one decision. The boundary rides the editor as the inline atom it already was, inside the control's `inline*` node, and the pairing pass looks inside the control, so a range that opens inside and closes outside keeps both halves instead of being deleted as an orphan. A revision covering a whole control that holds a marker still hoists to `w:ins > w:sdt`.

  The other range markers are still lifted: a `w:commentRangeStart` or a `w:moveFromRangeStart` inside the control is a marker the control does not own, and the pairing passes read it as a paragraph-level sibling.

- [#923](https://github.com/stella/folio/pull/923) [`20c9522`](https://github.com/stella/folio/commit/20c95224287ee8b7bfcb57defc17cac56a54fbcf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a bookmark marker in the container it was written in, so its range still covers what the author selected.

  `CT_Body`, `CT_Tc`, `CT_SdtContentBlock`, `CT_Row` and `CT_Tbl` each declare `w:bookmarkStart` and `w:bookmarkEnd` beside their own children, and Word writes them there whenever the selection was whole blocks, whole cells, whole rows or a whole table. folio re-anchored every one of them into a neighbouring paragraph. The element still reached the saved part, so nothing looked lost; what changed was the extent. A bookmark spanning a row came back inside one cell's paragraph, and a `REF` field or a link resolving it then covered the wrong text.

  `BlockContent` gains `BookmarkStart` and `BookmarkEnd` as members, so a marker on a body, a cell or a block content control is a block in its own right: it sits between the same two siblings in the model, in the ProseMirror document and in the saved part, with no index to keep honest. `TableRow` and `Table` gain `bookmarks`, a marker plus its position among the cells or rows, because neither models a child a marker could be. These stay typed rather than joining the verbatim sink: folio pairs a start with its end over the model, and a half kept as bytes leaves the other half unpaired and deleted on the first edit — which is the commoner shape, since a bookmark that opens on a row usually closes inside a cell.

  The editor gains a block-level `blockBookmarkBoundary` node, the block twin of the inline `bookmarkBoundary` atom, and a row's and a table's markers ride their node's attributes by reference the way an attribute remainder does. The boundary integrity pass reads all four carriers, so a pair spanning two levels stays whole.

- [#928](https://github.com/stella/folio/pull/928) [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Type `BorderSpec.style` as the `ST_Border` enumeration instead of `string`.

  `w:val` on `CT_Border` has 193 members. The model held it as a bare `string`, beside a hand-written 22-member `KnownBorderStyle` the parser narrowed against; the hand list omitted every page-border art glyph and four line styles, so a document that used one reached the consumers as a value no rendering table knew, and nothing could say which table was missing which member. `BorderStyle` is now generated from the committed schema graph by `bun run generate:border-styles`, and a schema refresh that adds a member fails the generator check rather than widening a `string`.

  `nil` and `none` stay distinct members: `none` cancels a border inherited from the container, `nil` states that none is set, and Word round-trips whichever the author wrote. Consumers ask `statesNoBorder`, `isBorderNone` or `isBorderNil` rather than comparing the token; `specifications/reserved-values` records the decision and the lint holds it.

  A `w:val` the schema does not declare is kept verbatim as `{ kind: "unrecognised", raw }` and written back unchanged, with a `border-style-outside-enum` parse warning so the normalisation is visible. Refusing it would drop an edge Word paints, and reading it as a default would rewrite the document on open.

  One table now says how a member renders. There were three — the layout bridge's, `formatToStyle`'s and `TableExtension`'s — and they covered different amounts of the enumeration, so a `thinThickSmallGap` cell edge came out `double` in the editor and `solid` on the paginated page, and a `dotDash` paragraph rule came out `dashed` through the bridge and `solid` through `borderToStyle`. `CSS_BORDER_STYLES` is total over the union at compile time, and a page border now takes the 3px floor for every member that paints as a CSS `double`, not only for `w:val="double"`.

  `KnownBorderStyle` is removed; `BorderStyle`, `BorderStyleValue` and `UnrecognisedBorderStyle` replace it. The layout engine's `BorderStyle.style` and `CellBorderSpec.style` are typed `CssBorderStyle`, which is what they always held, so a measurer or painter can no longer ask whether a laid-out border is `"nil"`. `TableCellBorderCommandSpec` and `TableBorderCommandSpec` are exported from `@stll/folio-core/prosemirror` and carry the same union, and both adapters' table-style presets use them instead of re-declaring the shape.

- [#931](https://github.com/stella/folio/pull/931) [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate `TableCellTextDirection` from `ST_TextDirection`. The union omitted
  `lrTb`, `lrTbV` and `tbLrV`, so a cell written with one lost the attribute at
  parse time and saved without it. The editor's writing-mode map is now total
  over the enumeration, and the display list reports an unpainted vertical flow
  for every direction the painter turns rather than the two that were listed.

- [#938](https://github.com/stella/folio/pull/938) [`9d1e897`](https://github.com/stella/folio/commit/9d1e8973631997ee478ca6ad0fedda4cc1134246) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read `w:sdtPr` through the shared child dispatcher, so a content control keeps every property its author wrote. `SdtProperties.preserved` holds the children folio does not model at their `CT_SdtPr` ordinal, and one writer serialises block, inline, row and cell controls from the model rather than replaying the source's bytes. `SdtProperties.rawPropertiesXml` is gone; `SdtProperties.lock` no longer reports `unlocked` for a value the reader refuses.

- [#927](https://github.com/stella/folio/pull/927) [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Describe a SmartArt diagram at parse time instead of rasterising it. `Image.preview` carries a bounded `PreviewDescriptor`, and the display list builds the PNG when it interns one, so a package's parse no longer pays megabytes per diagram for a picture nothing may paint. The rendered picture is unchanged.

- [#941](https://github.com/stella/folio/pull/941) [`0249b81`](https://github.com/stella/folio/commit/0249b811d88f80aa2d3d2bf2c0d1a9c20c8dc82d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rebuild `word/fontTable.xml` from the model without losing what the source declared.

  A repack copies the part across byte for byte, so the reader's gaps only showed on the paths that build a package from the model: a package folio authors, and a style set carried into a new document. `w:font` now goes through the shared child dispatcher, whose handler map the compiler makes total over the children `CT_Font` declares, so `w:notTrueType` lands in the ordered sink instead of on the floor and an attribute the font's record has no field for rides its remainder.

  Four things the model held or dropped are now written back. `w:charset` keeps the character set it names as well as the one it numbers, and a bare `<w:charset/>` — the default code page — is no longer written as no `w:charset` at all. The four `w:embed*` faces are written from the model with their `w:fontKey` and `w:subsetted`, which nothing wrote before although the relationship id was parsed: a rebuilt part pointed at no embedded font.

  `FontInfo.charset` becomes `FontCharset` and the four `embed*` fields become `EmbeddedFontRef`, so the key an embedded face cannot be decoded without travels with the relationship that names it. That retires the second font-table reader that existed only because the model dropped the key; one reader owns the part.

- [#941](https://github.com/stella/folio/pull/941) [`0249b81`](https://github.com/stella/folio/commit/0249b811d88f80aa2d3d2bf2c0d1a9c20c8dc82d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Rebuild `word/numbering.xml` from the model without losing what the part defines.

  A repack copies the part across and splices single definitions into it by id, so the reader's gaps only showed on the paths that build a package from the model. Measured over the cached public corpus, a rebuild used to lose 45 distinct element and attribute slots across the sampled packages that carry the part; it now loses none the part itself owns.

  `w:numbering`, `w:abstractNum`, `w:lvl`, `w:num` and `w:lvlOverride` go through the shared child dispatcher, whose handler map the compiler makes total over the children each content model declares. A `w:numPicBullet` and the `w:numIdMacAtCleanup` high-water mark land in the ordered sink in source position, and an attribute a record has no field for — `w15:restartNumberingAfterBreak`, `w15:durableId` — rides its remainder.

  `AbstractNumbering` gains `nsid` and `tmpl`, the identity Word recognises a list template by across documents; `ListLevel` gains `tplc`, `tentative`, `pStyle`, `lvlPicBulletId` and the `w:null` flag `w:lvlText` may carry. `lvlJc` widens to the whole `ST_Jc` enumeration `CT_Jc` declares, so a justification folio has no marker layout for is carried rather than taking the element with it; the three alignments layout does have are resolved from it. `NumberFormat` gains `bahtText` and `dollarText`, the two `ST_NumberFormat` members the model omitted.

  Three values came back different from the way they were written. `w:legacy` reads its own `w:legacy` attribute rather than a `w:val` the type does not declare, so an explicit "off" is no longer written as "on"; `w:legacySpace` and `w:legacyIndent` spelled with a unit resolve to the twips they count; an explicit `<w:isLgl w:val="0"/>` stays off. An empty `<w:pPr/>` or `<w:rPr/>` is written as the empty element the source wrote, not as an absent one, and a `w:lvl` whose `w:ilvl` names no level is kept as the definition it is while resolving to no level.

- [#931](https://github.com/stella/folio/pull/931) [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate `NumberFormat` from `ST_NumberFormat`. It omitted `bahtText`,
  `dollarText` and `custom`, and carried three `decimalZero{3,4,5}` members the
  format does not declare: the parser minted them from a custom format's pad
  width and the serializer wrote them back as a `w:val` no consumer can read.

  A custom format is now held as `custom` plus the `@w:format` it counts by
  (`ListLevel.numFmtFormat`), and written back as both. The three synthetic
  values move to `CounterFormat`, the render vocabulary `ListRendering.numFmt`
  and the editor's list attributes carry, which is never serialized.

- [#931](https://github.com/stella/folio/pull/931) [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate `ParagraphAlignment` from `ST_Jc` instead of spelling it by hand. The
  union omitted `start`, `end` and `numTab`, so a paragraph written with one
  parsed without an alignment and saved without a `w:jc`. `start` and `end` are
  direction-aware members, not spellings of `left` and `right`: the layout and
  the CSS projection resolve them against the paragraph's direction.

- [#936](https://github.com/stella/folio/pull/936) [`3fde3f4`](https://github.com/stella/folio/commit/3fde3f47fcdf038dd4a787a65fb46c4342752da1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every `w:pPr` child folio does not model, in the place the schema gives it, through a rebuild as well as through a replay.

  `CT_PPrBase` declares thirty-three properties and folio models fourteen. The rest reached disk only while the whole `w:pPr` was replayed as bytes, so the first edit to any paragraph property — an alignment command, a spacing change, a style applied — rebuilt the element without them: a table's `w:cnfStyle`, an East Asian document's `w:wordWrap` and `w:autoSpace*`, a frame's `w:mirrorIndents`, a vertical text box's `w:textDirection`. A style's property set was read by a second, narrower copy of the same `if`-chain, which knew neither `w:framePr` nor the Strict `w:ind` spellings.

  One reader now dispatches the set through the shared child dispatcher, so the compiler makes the handler map total over the declared children and each one carries a decision: modelled, kept as bytes, or named as another reader's. `w:rPr`, `w:sectPr` and `w:pPrChange` are the three with other owners. A handler that takes no typed value hands the child back instead of dropping it, which is what `<w:spacing/>` and a `w:jc` outside the reader's enumeration used to do.

  `ParagraphFormatting` gains `preserved`, the sink, recording each capture at its **schema ordinal** rather than at a count of modelled siblings: the count is a mirror of whichever properties folio models today, and it moves under the capture the moment one more of them is modelled. The order the four writers emit comes from a generated table in `@stll/docx-core/schema`, and one writer serves all four — a paragraph, a style, a numbering level, and the `CT_PPrBase` snapshot inside `w:pPrChange`.

  The cascade drops the sink rather than inheriting it: captured bytes belong to the element they were read from, and writing a style's back as direct formatting would outrank the tier they came from.

  A `w:pPrChange` whose original states nothing is no longer discarded for being empty. It records that the paragraph carried no direct formatting before the reviewer's edit, which is what rejecting the revision restores.

- [#927](https://github.com/stella/folio/pull/927) [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retire the preview rasteriser. `PreviewDescriptor` drops `pixelWidth` and `pixelHeight`, which sized a raster nothing builds any more, and `previewRaster.ts` goes with them; `MAX_PREVIEW_SHAPES` moves to the diagram reader, which is what applies it.

- [#936](https://github.com/stella/folio/pull/936) [`3fde3f4`](https://github.com/stella/folio/commit/3fde3f47fcdf038dd4a787a65fb46c4342752da1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the attributes a modelled property element carries that the model has no field for.

  A property element is an attribute bag, and the child dispatcher decides one whole: a handler either reads the element or hands back its bytes. `<w:ind w:leftChars="100"/>` survived because the reader took nothing from it, while `<w:ind w:left="720" w:leftChars="100"/>` — what a document actually carries — was modelled and lost the character unit. The same went for `w:spacing`'s line counts, `w:framePr`'s `w:hRule` and `w:anchorLock`, and the three attributes describing a `w:shd` pattern colour.

  The attribute remainder now rides the record that holds the element's modelled fields: `ShadingProperties`, `BorderSpec`, `TabStop` and `ParagraphFormatting.frame` gain `preservedAttributes`, and `w:ind` and `w:spacing`, which the model flattened into `ParagraphFormatting`, gain one remainder each there. The predicate is derived from the model rather than written beside each reader — `propertyElementAttributes.ts` holds one `as const satisfies` table per record — so a field added without an attribute to name does not compile.

  The container-survival census measured one attribute at a time and so could not see the defect at all. It now states each attribute pair a second time beside one the element models, and the pair survives only when it survives both.

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every `w:tblPr` and `w:sectPr` child folio does not turn into a typed value, in the order the content model declares.

  Both property sets were walked by a reader per property with no branch for the rest, so a child whose value the reader did not admit and a child folio models nothing for fell off the end: 19 table pairs and 16 section pairs in the survival census. The section serializer had a second answer to the same question, returning the empty string when the element carried any unread child, which failed the whole save rather than one property.

  Both now go through the shared child dispatcher with a handler map the compiler makes total over the schema's declared children, and a handler answers with what it took — a property the reader turned into nothing keeps its bytes. `TableFormatting.preserved` and `SectionProperties.preserved` hold them. `CT_TblPr` and `CT_SectPr` are sequences, so the sink records the schema ordinal rather than a count of modelled siblings, and both serializers merge modelled and captured children by it, reading the order from the generated declared-child list instead of restating it.

  `w:tblCaption`, `w:tblDescription`, `w:tblStyleRowBandSize` and `w:tblStyleColBandSize` are authored values rather than markup nobody reads, so `TableFormatting` models them as `caption`, `description`, `rowBandSize` and `columnBandSize`.

  A tracked property change no longer needs a non-empty snapshot to survive: `w:tblPrChange`, `w:trPrChange` and `w:tcPrChange` on a property set that states nothing of its own kept the author, the date and the id, and were dropped anyway.

- [#918](https://github.com/stella/folio/pull/918) [`3b984e5`](https://github.com/stella/folio/commit/3b984e5759f40fe5af1c658baa9c163078b7db69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `pic:cNvPr` and `wps:cNvPr`'s authored name and alt text when a drawing is rebuilt, instead of spelling the picture's name from the media filename and dropping the shape's.

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every declared child of a `w:trPr` and a `w:tcPr`, in the order the schema declares, through the editor as well as through a save.

  Both property sets were read by name: ten of the row's fifteen declared children had an `if` and the rest had nothing, so `w:cnfStyle`, `w:divId` and `w:tblCellSpacing` went on every save, as did `w:hMerge`, `w:headers` and the structural revision a `w:tcPrChange` snapshot records. So did any child whose value the reader refuses — a `w:trHeight` of zero, a `w:vAlign` the enumeration does not admit, an explicit off. Both sets now go through the shared child dispatcher, whose handler map the compiler makes total over the children the schema declares, and `TableRowFormatting.preserved` / `TableCellFormatting.preserved` hold what no reader took a typed value from.

  Each set is written by one call through `serializeSequenceChildren`, so the order is the generated declared-child list rather than the order of the serializer's statements: a `w:tcPr` whose children arrive out of `CT_TcPrBase`'s sequence comes back in it, and the sink's captures land between the same neighbours they were read between.

  An empty `<w:trPr/>` or `<w:tcPr/>` is kept. Both elements are optional, so a producer that wrote one stated something an absent element does not, and a parser that keyed the record on the properties the element yielded deleted it on save.

  `TableCellPropertyChange` gains `previousStructuralChange`: `CT_TcPrInner` declares the cell's insertion, deletion and merge, so a `w:tcPrChange` may record the cell as having stood inserted before the change, which is not the cell's current revision.

- [#938](https://github.com/stella/folio/pull/938) [`9d1e897`](https://github.com/stella/folio/commit/9d1e8973631997ee478ca6ad0fedda4cc1134246) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a row-level and a cell-level content control. A `w:sdt` between a table
  and its rows, or between a row and its cells, was unwrapped: the rows and cells
  were spliced in and the control — its tag, alias, lock, data binding and
  `w:sdtEndPr` — went on the floor, so a template whose repeating section or
  bound cell folio merely opened and saved came back unbound.

  `TableRow.contentControls` and `TableCell.contentControls` record the control
  on each child it wrapped, outermost first, and the save re-opens one wrapper
  per run of consecutive children that name the same control. The record is on
  the children rather than between them because a table's children are rows and a
  row's are cells, and neither has a node to spare for a wrapper that is not one;
  it rides the ProseMirror row and cell nodes as an attr, so splitting or moving
  one keeps it inside its control. A control over several rows, and a control
  inside a control, both come back as they were written.

  `SdtProperties.endProperties` models `w:sdtEndPr`, which folio held only as
  captured bytes: every control it rebuilt — one an edit touched, one a full
  repack wrote — lost its end mark and the run properties on it. This covers the
  block and inline levels too.

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the table properties a row overrides (`w:tblPrEx`), and the revision that records changing them.

  Nothing read the element. The row's child walk called it `OWNED_ELSEWHERE` — "another reader owns this" — and no reader did, so the nine table properties a row may restate and the `w:tblPrExChange` beside them went on every save: 24 pairs in the survival census, from the row pair down through `CT_TblPrEx`, `CT_TblPrExBase` and `CT_TblPrExChange`. Word writes the element when a table is built by merging two, and a consumer reads it in place of the table's own properties for that row, so the loss restyled the row.

  `TableRow.tablePropertyExceptions` holds it as the same `TableFormatting` the table carries, because `CT_TblPrEx` is the middle of `CT_TblPrBase` and the two are read by one set of handlers rather than two. `TableRow.tablePropertyExceptionChanges` holds the revision, as `Table.propertyChanges` holds `w:tblPrChange`. Both ride the ProseMirror row node, so the element survives the editor as well as a save.

  `CT_Row` declares `w:tblPrEx` before `w:trPr`, and the serializer writes it there. The element is optional, so an empty one is kept rather than read as no exceptions at all: its presence is the value.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every `w:rPr` child folio does not model, where it stood, for a run, a paragraph mark, a style and a tracked property change alike.

  A run property set was walked by a reader per property with no branch for the rest. `w:bdr`, `w:fitText`, `w:eastAsianLayout`, `w:snapToGrid`, `w:webHidden`, `w:specVanish` and `w:oMath` had no model at all; `w:rFonts`, `w:u`, `w:lang`, `w:w` and `w:sz` were read and dropped whenever the reader took no typed value from them; and the paragraph mark's whole `w:rPrChange`, along with everything inside it, went with them. A save that rewrote the element — which is every save after an edit — lost all of it.

  `EG_RPrBase` now goes through the shared child dispatcher, with a handler map the compiler makes total over the children the schema declares. A handler answers with what it took, so a property the reader turned into no typed value keeps its bytes: a name-keyed map can state the names folio has never heard of, not the values a reader refuses. `TextFormatting` gains `preserved`, and because it is a sequence the sink records each capture's schema ordinal rather than a count of modelled siblings.

  The four owners of a run property set — a run, the paragraph mark inside `w:pPr`, and the snapshot inside either one's `w:rPrChange` — share that map and differ only in which children a sibling record has already claimed, which the call site names. One writer serves all of them plus a style and a numbering level, and it orders its children from the generated declared-child list rather than from the order of its own statements: folio wrote `w:vanish` before `w:noProof` while the schema declares the reverse, which a validating consumer refuses.

  Captured bytes belong to the element that was parsed and to no other, so style resolution and formatting merges drop them rather than inheriting them onto every run below.

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a `w:smartTag` and a run-level `w:customXml` as the wrappers they are.
  folio spliced a smart tag's children into the paragraph and kept no wrapper, so
  the tag, its namespace and its properties were gone on the first save; and it
  captured a run-level `w:customXml` whole, so the wrapper came back but every run
  inside it was opaque bytes the editor could not touch. Both are `InlineWrapper`
  kinds now — `smartTag` and `customXml`, each carrying `element`, an optional
  `uri` and the `w:smartTagPr` / `w:customXmlPr` verbatim — so their content is
  parsed by the same run-level walk `w:bdo` and `w:dir` take, nests with them in
  either order, and comes back through the editor on the same `inlineWrapper`
  mark. The properties are part of the mark's stack key, so two adjacent tags that
  differ only in their properties stay two tags; they are replayed only when they
  are structurally the element they claim to be, because a paste from outside the
  editor can put any string on a mark.

- [#931](https://github.com/stella/folio/pull/931) [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate `TabStopAlignment` from `ST_TabJc`. The union omitted `start` and
  `end`, and a stop declared with either left the model entirely, because the
  reader needs both a position and an alignment to keep one. The numbering
  parser's second, hand-rolled reader for the same enumeration read both as
  `left`; it now narrows against the one picklist.

- [#932](https://github.com/stella/folio/pull/932) [`166d3f0`](https://github.com/stella/folio/commit/166d3f0868393dc74c0609ac4a94cba0e579743b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read `w:tblPr/w:jc` and `w:trPr/w:jc` against `ST_JcTable`. The two had a
  reader of their own that accepted `left`, `center` and `right`, folded `start`
  onto `left` at the table and refused it at the row, so a table written `start`
  saved as `left` and one written `end` saved with no `w:jc` at all. `start` and
  `end` are members now, resolved against the table's `w:bidiVisual` at layout
  and written back as authored.

- [#932](https://github.com/stella/folio/pull/932) [`166d3f0`](https://github.com/stella/folio/commit/166d3f0868393dc74c0609ac4a94cba0e579743b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Pair `ST_TextDirection`'s two spellings of each flow by ECMA-376 Part 4
  §14.11.7. Folio paired them by the letters in the token, so every one of the
  six Strict spellings rendered as something other than its Transitional twin:
  `tb` is the horizontal flow and turned a quarter clockwise, `rl` and `lr` are
  vertical flows and painted flat. Rendering is now decided per flow, and the
  section's own text direction is narrowed against the enumeration rather than a
  second hand-written copy of it.

- [#930](https://github.com/stella/folio/pull/930) [`56539e3`](https://github.com/stella/folio/commit/56539e3504cf2ed26e6b2e016bd66507581d5c24) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Replace `ThemeColorSlot` with `ThemeColor`, generated from `ST_ThemeColor`, and
  carry a token outside it as `{kind: "unrecognised", raw}` rather than dropping
  it. `ColorValue.themeColor` now holds either; `themeColorSlot` resolves one to a
  theme slot, through the `w:clrSchemeMapping` key for the mapped members.

- [#922](https://github.com/stella/folio/pull/922) [`d098792`](https://github.com/stella/folio/commit/d098792ef0634154446a13d80757b7f73b233838) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Model `w:tblGridChange` as the grid it records rather than as the bytes it arrived in.

  A `w:tblGridChange` holds a `w:tblGrid` of its own, and that grid holds its own `w:gridCol` children: a container nested in one of its own kind. It travelled as `TableFormatting.gridChangeXml`, a verbatim slot, so the snapshot's grid and every column in it existed only as markup nothing could read, and a rebuild could only copy the string back. The public corpus has the element in 28 packages, 104 columns in all, so this is a shape documents actually carry.

  `TableFormatting.gridChangeXml` is replaced by `TableFormatting.gridChange`, a `TableGridChange` holding the revision's `@w:id` and one entry per `w:gridCol`. `w:w` is optional on a `w:gridCol`, so a column the snapshot stated no width for is `undefined` rather than zero, and it is written back without a width: a snapshot is a record of what stood, and a column with no measure is not a column of width zero.

- [#919](https://github.com/stella/folio/pull/919) [`fa2abc1`](https://github.com/stella/folio/commit/fa2abc134692faf7dd48cdeb28d0631ee9a796b7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Model `wp:wrapPolygon` on `ImageWrap` and write the authored outline back, instead of a constant 21600-unit rectangle for every tight and through wrap. Wrap insets now record whether `wp:inline`/`wp:anchor` or the `wp:wrap*` child stated them, so a rebuild writes each one where it was authored.

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a transparent wrapper inside a link or a simple field as the wrapper it
  is. `CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`, which declares
  `w:bdo`, `w:dir`, `w:smartTag` and the run-level `w:customXml`, so
  `w:hyperlink > w:bdo > w:r` is markup a producer may write; folio captured all
  four whole, which kept the markup and made every run inside it opaque bytes the
  editor could not touch. `Hyperlink["children"]` and `SimpleField["content"]`
  now carry `InlineWrapper`, the wrapper's children are walked by the container's
  own handler map — so a `w:ins` inside a `w:bdo` inside a link is captured for
  the same reason a `w:ins` directly inside the link is — and the runs reach the
  editor carrying both the link mark and the `inlineWrapper` stack. Saving from
  the editor writes the canonical order the wrapper design fixed, revision then
  wrapper then hyperlink then run, so a link authored inside a `w:bdo` comes back
  with the `w:bdo` around it; the link, the wrapper and the text all survive, and
  a paragraph nobody edited keeps its authored order because selective save
  replays its bytes. A simple field keeps the wrapper inside itself, because its
  node holds its own inline content.

- [#917](https://github.com/stella/folio/pull/917) [`8f9a01b`](https://github.com/stella/folio/commit/8f9a01b2b849e743a3dbac019e7a1aecbf9c2379) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `wp:anchor`'s `simplePos`, `relativeHeight`, `locked` and `hidden`, the `wp:simplePos` offsets, and both `wp:docPr` links whole when an edited drawing is rebuilt from the model.

- [#940](https://github.com/stella/folio/pull/940) [`f1a4d2d`](https://github.com/stella/folio/commit/f1a4d2dd55fc83bc3872253fd4b00a017785ec85) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write an explicit off for every `CT_OnOff` element. `serializeOnOffElement` is the
  one writer: absent writes nothing, an on writes the bare element, and an off
  writes `w:val="0"`, which is what cancels an inherited on. The row, cell, table,
  control and paragraph-mark readers keep the three states apart as well, and a
  control that states nothing keeps stating nothing through the editor.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Model `w:outlineLvl` as a union, so the reserved body-text value cannot be read as a tenth heading level.

  `ParagraphFormatting.outlineLevel` was `number`, which put ECMA-376 17.3.1.20's rule ("9 specifically indicates that there is no outline level applied to this paragraph") in every consumer's hands. `OutlineLevel` is now `{ kind: "bodyText" } | { kind: "heading"; level: 0..8 }`, with `level` a union of nine literal types: the sentinel has no representation as a heading, an out-of-range value has none at all, and an absent field still means "states none, inherits one".

  One reader owns the parse boundary (`outlineLevelFromStatedValue`) and one writer owns the emit (`outlineLevelStatedValue`). The paragraph parser, the style parser, the style cascade, the display-list outline, the layout bridge, the ProseMirror attr and its validator, markdown, the style sets and the legal-source compiler all move to the union; `isHeadingOutlineLevel` and the bare `BODY_TEXT_OUTLINE_LEVEL = 9` are gone, replaced by `headingLevelOf` and the body-text arm.

  A `w:outlineLvl` outside 0..9 is now dropped at the parse boundary rather than carried through the model, which is what the Rust projection kernel already did. The container-survival census records the one value that stops surviving a rebuild.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Model `w:numPr` as a union

  `ParagraphFormatting.numPr` and `numPrFromStyle` carry
  `ParagraphNumberingOverride` instead of two optional slots, so the reserved
  `w:numId w:val="0"` has no representation past the parse boundary and a level
  stated without an id is a named arm rather than a half-filled pair. The
  cascade fold `mergeParagraphNumbering` replaces the object spreads that used
  to restate ECMA-376 17.3.1.19 at each tier.

- [#918](https://github.com/stella/folio/pull/918) [`3b984e5`](https://github.com/stella/folio/commit/3b984e5759f40fe5af1c658baa9c163078b7db69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read and write `wp:anchor`'s own attributes from one `DrawingAnchor` record shared by pictures, shapes and text boxes, and carry it through the editor.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give `w:numPr` one union, one cascade fold and one reader.

  `ParagraphNumberingOverride` is what a tier states — `none` (the reserved `w:numId 0`), `reference` (an id and an optional level) or `levelOnly` — and `ResolvedParagraphNumbering` is what the cascade leaves. Three arms, not two: `w:numId` and `w:ilvl` inherit independently (ECMA-376 17.3.1.19), so a tier that states only the level keeps the id it inherits, and that shape is what Word writes whenever a styled list paragraph is demoted.

  `mergeParagraphNumbering` is that inheritance written once. It replaces the paragraph parser's object spread, and it is closed under itself and associative over the three cascade tiers, so a third tier needs no special case.

  `paragraphNumberingFromSlots` is the one mapping from the element's two slots onto an arm, and `readParagraphNumbering` reads the element. Both are exported from `@stll/folio-core/docx` alongside `NO_NUMBERING_NUM_ID` and `isNumberingReference`, which now live in `@stll/docx-core` where the model does. Three duplicate spellings of the reserved id are retired: the hand-inlined copy in `docx-core`'s validator, the bare literal in the operation reader, and the relational form in the AI snapshot, which was the one spelling that read a malformed package's negative id as "not numbered" while every other spelling read it as a dangling reference.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write a compiled `w:rPr`'s children in the schema's order, from the same generated list folio-core writes from.

  `@stll/docx-core` holds a second `w:rPr` writer — the one the legal-source compiler and the build-from-scratch export share — and it had grown an order of its own, emitting `w:highlight`, `w:sz` and `w:szCs` ahead of `w:rFonts`. A run carrying both a font and a size therefore came out in one order from this package and another from folio-core's serializer. `EG_RPrBase` is an `xsd:choice` referenced `maxOccurs="unbounded"`, so both spellings are valid; what the canonical order buys is one form, the one Word writes, from both writers.

  The generated order moves down to where both can read it: `@stll/docx-core/schema` is a new subpath exporting `SEQUENCE_CHILDREN` and the writer that orders by it, and folio-core's declared-child table spreads that same object in. One emitted order, one sort, and a serializer that cannot restate either.

- [#928](https://github.com/stella/folio/pull/928) [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Type a shape or text-box outline's dash as `ST_PresetLineDashVal`, and give each stroke vocabulary its own table.

  `ShapeOutline.style` was a hand-written eleven-member union named after CSS, holding what `a:ln/a:prstDash@val` declares. It is now `ShapeOutline.dash`, typed with `PresetLineDashVal`, generated from the committed schema graph by `scripts/generate-preset-line-dash.ts` with the same write/check pair `BorderStyle` uses; `bun run generate:preset-line-dash:check` runs in CI. A `@val` the schema does not declare is kept as `{ kind: "unrecognised", raw }`, written back unchanged, and reported through `ParseContext` as `outline-dash-outside-enum`. `a:custDash` is a different element and stays unmodelled: an outline that carries one replays through `ShapeOutline.rawXml`.

  The display list resolved three vocabularies through one lookup keyed by lower-cased strings: a CSS `border-style`, a DrawingML preset dash, and a CSS `text-decoration-style`. `dash`, `dot` and `solid` collide across them, and every member no other vocabulary spells the same way had no entry and painted as a plain line. Nine of the eleven preset dashes (`dot`, `lgDash`, `dashDot`, `lgDashDot`, `lgDashDotDot`, `sysDash`, `sysDot`, `sysDashDot`, `sysDashDotDot`) and the seven heavy underline members were in that set, so a `sysDash` outline and a `dottedHeavy` underline both stroked solid. There are now three tables, each `as const satisfies Record<Union, StrokePattern>` over its own vocabulary, and each consumer calls the one it speaks.

  The DOM painter had the same defect one step further on: it interpolated the outline's dash straight into a CSS `border` shorthand, so `border: 2px sysDash #000` was invalid and a dashed text-box outline did not paint at all. A dash is now translated to a CSS keyword before it reaches a shorthand, and `run.underline.style` is translated rather than assigned, which is what made `text-decoration-style: dottedHeavy` a no-op.

- [#925](https://github.com/stella/folio/pull/925) [`da1fc6a`](https://github.com/stella/folio/commit/da1fc6a4c8d7a48705187d164d8e8180cefeff74) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write a `wp:effectExtent` back on the element that authored it, so a wrap child's own reservation survives a rebuild.

  `CT_Inline` and `CT_Anchor` declare a `wp:effectExtent`, and so do `CT_WrapSquare` and `CT_WrapTopBottom`. They are two values: the drawing's is the object's own effect reservation, the wrap child's is the reservation the text flow is computed against. folio read only the drawing's, into `Image.padding`, and wrote it back there, so a wrap child's own reservation round-tripped an untouched document on the strength of its captured bytes and was gone the moment anything forced the serializer.

  `ImageWrap` gains `effectExtentSlots`, the `distanceSlots` shape one element over: `drawing` and `wrapChild`, each holding the element's four sides. The value in force stays where its consumers read it, on `Image.padding`. `resolveEffectExtents` decides the rebuild the way `resolveWrapDistances` decides the insets — each reservation goes back on the element that stated it while the drawing's is unmoved, and once an editor has resized it the rebuild states the value in force on the drawing alone rather than keeping a wrap reservation computed against a shape that is no longer there. The slots ride through the editor as `wrapEffectExtentSlots` on the image, shape and text-box nodes.

  A shape and a text box have never had a reservation of their own — the rebuild wrote `l="0" t="0" r="0" b="0"` on every one of them — and now keep the one they were authored with.

### Patch Changes

- [#931](https://github.com/stella/folio/pull/931) [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record `FontHint`'s `cs` as a deliberate widening of `ST_Hint` rather than a
  divergence waiting to be closed. The registry gains a `wider-than-schema`
  verdict that carries the citation and may not grow without one.

- [#941](https://github.com/stella/folio/pull/941) [`0249b81`](https://github.com/stella/folio/commit/0249b811d88f80aa2d3d2bf2c0d1a9c20c8dc82d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record the reserved-value decision for every field the model declares.

  `ImageWrap.distanceSlots`, `ImageWrap.polygon` and the font table's two verbatim sinks landed without an entry in the registry their types are total over, so `typecheck:reserved-values` did not compile. Each carries markup rather than a value with a reserved meaning, and now says so.

- [#927](https://github.com/stella/folio/pull/927) [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give a drawing that carries no relationship one spelling, and classify the preview the second spelling hid. `Image.rId` is a `RelationshipId`, a branded non-empty string the parser mints from a real `r:embed`, `r:id` or `r:link`, so the empty string can no longer stand for absence: it reached a save as `<a:blip r:embed=""/>`. A VML shape's render is now preview-only, like the group render beside it, so the editor declines to manipulate it and a save replays the authored `w:pict` instead of writing the render into `word/media/` as the picture the shape had become. Stored collaboration snapshots take both changes through attr-schema version 4.

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve `w:tblPrExChange` like every other tracked property revision, from one list of them.

  The element round-tripped but the editor did not know it: accept, reject and the tracked-change list each carried their own list of four or five change elements, and none of them named the fifth. Accepting every change left the revision on the row, so the document said a formatting change was still pending after the reviewer had resolved it.

  The set is now written down once. `PROPERTY_REVISION_KINDS` is the model's census of the change elements that store a complete previous property set, and one site table says where each one lives, how it resolves and what a reader calls it. The carrier reader, the accept/reject command, the tracked-change list, the comparison's scopes and the Vue sidebar's labels are each total over it, so a revision the model gains is a compile error at every one of those rather than a branch nobody wrote.

  Two revisions the list had already lost come back with it: a paragraph's `w:pPrChange` was read from an attr the schema does not declare, and `w:sectPrChange` was never listed at all.

  Accepting a `w:tblPrExChange` drops the record and keeps the row's current exceptions; rejecting it restores the stored ones wholesale, including restoring their absence.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop writing `<w:ilvl w:val="undefined"/>` for a paragraph that stated no level.

  `serializeDocumentToDocx` wrote `w:ilvl` unconditionally beside `w:numId`, so a `numPr` carrying only an id produced an attribute value `CT_DecimalNumber` does not accept. An absent `w:ilvl` is level zero and is not the same bytes as a stated `w:val="0"`, so it stays absent.

## 0.24.0

### Minor Changes

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every block-level child folio does not model, where it stood, through the editor as well as through a save.

  `w:body`, `w:hdr`, `w:ftr`, `w:tc`, an SDT's content and a note body share one walk, and it modelled paragraphs, tables and content controls and let the rest fall off the end. A `w:permStart` between two paragraphs is the whole of a document-protection range; `w:altChunk` is an entire imported document; `m:oMathPara` is a display equation. The walk now goes through the shared child dispatcher, whose handler map the compiler makes total over the children the schema declares for a block container and whose default is the verbatim sink.

  `BlockContent` gains a `preservedBlock` member holding the captured markup, and the editor gains a zero-width `preservedBlock` node for it. Position is structural on both sides: the capture sits between the same two blocks in the model, in the ProseMirror document and in the saved part, so inserting, splitting or deleting a neighbour moves it the way a reader would expect and nothing has to keep an index honest.

  `Paragraph`, `Table` and `BlockSdt` lose `rawMarkersBefore` / `rawMarkersAfter`, the narrower mechanism this replaces: it kept only sixteen range-marker names, dropped them when the container held no block at all, and had no editor leg, so a document that survived an untouched save lost the markup the moment anybody opened it. `Footnote.content`, `Endnote.content` and `TableCell.content` are now `BlockContent[]` rather than hand-written copies of it.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every child a `w:hyperlink` or a `w:fldSimple` holds and folio does not model, where it stood, through the editor as well as through a save.

  `CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`: either may hold a permission range, a proofing error, a transparent wrapper or one of the eight custom-XML revision ranges between its runs. The link parser modelled the run and the two bookmark boundaries and returned `null` for everything else; the field parser read `w:r` and `w:hyperlink` and skipped the rest. Both walks now go through the shared child dispatcher, over generated declared-child sets the compiler makes their handler maps total over, with the verbatim sink as the default.

  `Hyperlink["children"]` and `SimpleField["content"]` gain `PreservedInline`, so the capture is a member of the container's own content union and stands between the same two children in the model, in the ProseMirror document and in the saved part. The editor carries it as the opaque atom the paragraph level already uses, inside the link mark, so it moves, is accepted and is rejected with the link. A simple field holding one keeps its children rather than collapsing to its display text, which would have dropped the markup on the way out of the editor.

  The link's handler map is exported and read twice: by the link parser, and by the paragraph parser's revision-segmenting walk, which overrides only the four `CT_RunTrackChange` wrappers because OOXML nests a revision inside a link and the model nests the link inside the revision. A child one of them starts recognising is recognised by both.

- [#906](https://github.com/stella/folio/pull/906) [`34ee586`](https://github.com/stella/folio/commit/34ee58684fd0700088e92deae7900198878c6849) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the transparent wrapper a tracked change was authored around inside the
  revision. `TrackedRunContent` held no `w:bdo`/`w:dir` and no inline `w:sdt`, so
  the parser lifted one out to a sibling and
  `<w:ins><w:bdo>x</w:bdo></w:ins>` saved as `<w:ins/><w:bdo>x</w:bdo>`: `x` was
  no longer inserted, and accepting the revision kept it exactly as rejecting it
  did. `TrackedRunContent` and `InlineSdt["content"]` now admit both wrappers, a
  single admission map bound to those content types decides what each wrapper
  keeps, and the serializer carries the revision's disposition through the
  wrapper so a `w:del` still writes `w:delText` around it. The opposite authored
  order, a revision inside the wrapper, is unchanged.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the children a table row holds beside its cells, between the same two cells, through a save.

  `CT_Row` declares a permission range, a proofing error, the row-level comment and move ranges and the eight custom-XML revision ranges beside `w:tc`. The row walk read `w:tc`, unwrapped `w:sdt` and carried a bookmark boundary into a neighbouring cell's paragraph; everything else it returned from. The walk now goes through the shared child dispatcher over a generated `row-content` set the compiler makes its handler map total over, with the verbatim sink as the default and `w:trPr` / `w:tblPrEx` marked as read elsewhere so neither is written twice.

  `TableRow` gains `preserved`, the ordered sink whose `index` counts the cells that preceded a capture. A row-level child cannot be a cell, so this is the sink case rather than the union case the inline levels use.

  The editor leg stops at the save: the table schema has no row-level node a zero-width capture could be, and an index recorded on the row node would drift the first time a column moved. The contract records those pairs as `editorProjection` rather than `neverParsed`, which is the difference between markup the model holds and markup folio never read.

- [#901](https://github.com/stella/folio/pull/901) [`051dbd6`](https://github.com/stella/folio/commit/051dbd612dc6541df1725a29d7bfea8612bed056) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the children a table holds beside its rows, between the same two rows, through a save.

  `CT_Tbl` declares a permission range, a proofing error, the table-level comment and move ranges and the eight custom-XML revision ranges beside `w:tr`. The table walk read `w:tr` and unwrapped `w:sdt`; everything else it returned from. The walk now goes through the shared child dispatcher over a generated `table-content` set the compiler makes its handler map total over, with the verbatim sink as the default and `w:tblPr` / `w:tblGrid` marked as read elsewhere so neither is written twice. A `w:customXml` row wrapper is kept whole rather than dropped, and a `w:tbl` nested directly in a `w:tbl` reaches the sink through its default rather than being flattened into the rows around it.

  `Table` gains `preserved`, the ordered sink whose `index` counts the rows that preceded a capture. A table-level child cannot be a row, so this is the sink case rather than the union case the inline levels use.

  The editor leg stops at the save, as it does for the row sink one level down: the table node's children are rows, and a zero-width capture between two of them is not a row.

- [#909](https://github.com/stella/folio/pull/909) [`817c3cf`](https://github.com/stella/folio/commit/817c3cf4ed5399571dd88bbc0dbc3ce4313c5068) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry a transparent inline wrapper into the editor instead of dropping what it said.

  `w:bdo` and `w:dir` reached the editor as their content and nothing else: the projection flattened them, so opening a document lost the direction the author wrote and the painter drew the text in the paragraph's direction. The tree is still flattened — the inline loops narrow by a chain of `else if`, and a wrapper left in one would reach whichever branch happens to be last — but what the wrapper said now rides an `inlineWrapper` mark on the leaves it held. The mark's `stack` attr lists the wrappers a leaf sits inside, outermost first, because ProseMirror's mark set is unordered across types and two marks could not say which wrapper is inside which. `RunFormatting` gains `bidiWrapper`, the painter writes `unicode-bidi` and `dir` from it, and a glyph run takes the wrapper's direction over the paragraph's.

  `BidiWrapper` becomes `InlineWrapper`, discriminated on `kind`, with `type: "inlineWrapper"`. The old name admitted only one kind of transparent wrapper; a smart tag and a custom-XML wrapper are the same shape and become added members rather than new types every exhaustive switch has to learn. Only `bidi` exists today: nothing parses the other two yet.

  `AUTOSAVE_FORMAT_VERSION` moves to 3, because the codec serialises the model and an envelope written under 2 holds paragraph content under the old discriminator. A stored collaboration snapshot is unaffected and the attr-schema version does not move: the new mark attr defaults to `null`, which is what every existing snapshot means.

  The save leg is unchanged. An edited wrapper span still loses its wrapper, because `fromProseDoc` rebuilds the wrapper from the source paragraph rather than from the mark.

- [#901](https://github.com/stella/folio/pull/901) [`051dbd6`](https://github.com/stella/folio/commit/051dbd612dc6541df1725a29d7bfea8612bed056) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the attributes an element carried that folio has no field for, through a save and through the editor.

  Word writes a revision-session id on nearly every paragraph, run, row and section (`w:rsidR`, `w:rsidRPr`, `w:rsidDel`, `w:rsidP`, `w:rsidRDefault`, `w:rsidTr`, `w:rsidSect`). folio rebuilt each of those elements from the model alone, so opening a document and saving it rewrote the whole revision history.

  `Paragraph`, `Run`, `TableRow` and `SectionProperties` gain `preservedAttributes`, an ordered list of resolved `{ namespace?, name, value }` records. The decision of what to keep is made on the resolved namespace URI and local name, so a source that binds a second prefix to the WordprocessingML namespace does not get a second copy of an attribute the parser already read; the writer is handed the modelled attributes it is about to emit and drops any remainder entry that would spell one of them again, so a duplicate attribute cannot reach the part. A namespace declaration is never in the remainder, and neither is an attribute whose namespace the rebuilt part cannot bind.

  The remainder follows the record: a paragraph, row or section the editor creates from scratch has none, an authored one's survives `toProseDoc`/`fromProseDoc` unchanged, and when a command splits a record in two the half that comes first in document order keeps it. A run has no record in the editor — it is text plus marks — so a run's remainder survives a save and not the projection.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every inline child folio does not model, at its source position. One walk serves a paragraph, the four run-level tracked-change wrappers, `w:bdo`/`w:dir` and an inline content control, and it now goes through the shared child dispatcher over a handler map the compiler makes total. `ParagraphContent` gains a `preservedInline` member holding the captured markup, so `w:permStart`, `w:proofErr`, `w:customXml`, the eight custom-XML revision ranges and `w:subDoc` survive a save and the editor round trip.

  Inside a tracked change the position is the point: markup lifted out of a `w:ins` is markup the reviewer no longer accepts or rejects with the change, so the capture sits inside the wrapper in the model, in the editor and in the saved part. `w:customXml` also keeps the text it puts on the line.

  A bare OMML element is now read by namespace rather than by falling off the end of a switch, and `m:oMathPara` keeps its display form.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every run child folio does not model instead of letting it fall off the end of the run-content switch. `RunContent` gains a `preservedXml` member holding the captured markup at its source position, plus the visible text it contributes, so `w:ruby`, `w:contentPart`, `w:pgNum`, `w:annotationRef`, the note markers and any foreign or future element survive a save and read as text.

  The keep rule now asks the model rather than the source element. The two disagreeing was a two-save oscillation rather than a loss: the first save wrote a run whose payload the model never held, the next parse dropped that run, and the second save differed from the first.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove the verbatim sink's attribute remainder. `PreservedMarkup.attributes`, the `PreservedAttribute` type, the dispatcher's `modelsAttribute` option and `serializePreservedAttributes` had no caller in the product: no container ever passed the predicate, so no attribute was ever kept, and the shape read as coverage that was not there. `docs/container-contract.md` records the design and what wiring it needs.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add the ordered verbatim sink and the shared child dispatcher, and put `w:comment` bodies on them.

  `PreservedMarkup` holds a container's unmodelled children with their position relative to its modelled ones, plus an ordered attribute remainder, so the serializer puts them back between the same siblings rather than at the end. `dispatchChildren` walks a container with a handler map the compiler makes total over the children the schema declares for it, and routes anything undeclared — a foreign namespace, an `mc:` construct, an element a later OOXML revision adds — to the sink by default.

  A comment body may hold everything a document body can. folio modelled only `w:p`, so a table, an equation, a content control, a bookmark or a range marker in a reviewer's comment disappeared on save; `Comment.preserved` now keeps them.

### Patch Changes

- [#892](https://github.com/stella/folio/pull/892) [`9035639`](https://github.com/stella/folio/commit/90356394a8d68e078bbaa95ad2e6643fcff51a1f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `@w:fldLock` and `@w:dirty` as a field authored them, on `w:fldSimple` and on the `w:fldChar` that opens a complex field. Three readers and three writers had each collapsed the two attributes to "present and true", so an explicit `w:dirty="0"` -- a field inside a `TOC` result that says not to recompute -- parsed as an absence and saved as one. Both directions now live in one module, `docx/fieldState`. The editor keeps the distinction too: the field node's `fldLock` and `dirty` attrs default to absent rather than `false`, so projecting a field through the editor no longer invents an explicit off. The attributes are written as `1`/`0`, matching Word. Because the persisted attr shape changes, the collaboration attr schema goes to version 2, and `migrateFolioYjsSnapshot` drops the `false` a version-1 snapshot stored for a field that authored neither flag.

- [#888](https://github.com/stella/folio/pull/888) [`4f8eed6`](https://github.com/stella/folio/commit/4f8eed643be10173a177695b343e36e2d02d98cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open every document that carries an explicit page-break run. `w:br w:type="page"` is an ordinary run child, so Word writes one in a bordered, framed or outlined paragraph, inside a table cell or a text box, and beside any inline kind. Folio refused several of those shapes at conversion and again at layout, which meant the document could not be opened in the editor, laid out or exported to PDF at all. They now project, save and round-trip; where layout can only approximate the break's owner, it says so through the parse-warning channel under the new `page-break-projection-approximated` code instead of throwing. `UnsupportedDocxToProseMirrorConversionError` goes with the last refusal that raised it.

## 0.23.0

### Minor Changes

- [#881](https://github.com/stella/folio/pull/881) [`05044c5`](https://github.com/stella/folio/commit/05044c53a4b02600a661c834cb65b09d5f31a56f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read and write a decorative image as the extension Word writes, and keep `hidden` a separate fact. `Image.decorative` was read from a `@decorative` attribute `CT_NonVisualDrawingProps` does not have (no file in the public corpus writes one), and written back as `hidden="1"`, which says the drawing is not displayed — so a decorative image became a hidden one, and re-parsed as neither. It now round-trips through `wp:docPr`'s `{C183D7F6-B498-43B3-948B-1728B52AA6E4}` extension, `Image.hidden` carries `@hidden` on its own and is written identically for inline and anchored drawings, and `Image.docPrExtensions` keeps the other `a:ext` entries of the same list verbatim and in order rather than dropping them. All three survive the editor round trip.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop minting `wp:docPr@name`. A shape, text box or picture whose name the model
  did not carry was written back as `Shape 3`, `TextBox 3` or `Picture 3`, so a
  connector named `直接箭头连接符 2` came back in English through the editor round
  trip and no later reader could tell a generated name from an authored one. The
  serializer now writes only what the model holds, and the insert command names
  the object it creates; a drawing with no name writes `@name=""`, the required
  attribute with nothing in it.

  `wp:docPr@descr` (alt text) and `@title` were never modelled for shapes and text
  boxes at all, so a rebuild dropped them: `Shape` and `TextBox` gain `alt` and
  `title`, the ProseMirror shape and text-box nodes carry them along with the
  authored name, and one reader/writer pair owns all three attributes for every
  drawing kind.

- [#881](https://github.com/stella/folio/pull/881) [`05044c5`](https://github.com/stella/folio/commit/05044c53a4b02600a661c834cb65b09d5f31a56f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give XML escaping one owner, and make its output always well-formed. `@stll/docx-core` now exports `escapeXmlText` and `escapeXmlAttribute`: six hand-rolled escapers disagreed about the characters that matter, so a value could leave folio as markup Word refuses to open, or come back changed. Both functions drop the characters XML 1.0 §2.2 forbids (the C0 controls outside tab/LF/CR, U+FFFE, U+FFFF, unpaired surrogates), which cannot be escaped into a document either. The attribute form writes tab, LF and CR as character references, because §3.3.3 has every conformant reader flatten a literal one to a space; the text form does the same for CR, which §2.11 would otherwise rewrite to LF. `sanitizeXmlCharacters` applies the same rule at an input boundary, where the value can still be reported.

## 0.22.0

### Minor Changes

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `w:bdo` and `w:dir`, the Unicode bidirectional controls. `w:dir` is an embedding and `w:bdo` an override, and folio discarded both on save — in a right-to-left document that is the difference between a readable line and a scrambled one, because an override is what makes a Latin word inside it read backwards. They are now a `BidiWrapper` member of `ParagraphContent`, a transparent inline container that nests, holds anything paragraph content holds, and that every paragraph walk reads straight through. The editor projection flattens the wrapper for now, so a document edited in the editor still loses the direction; the save path keeps it.

- [#873](https://github.com/stella/folio/pull/873) [`48715e3`](https://github.com/stella/folio/commit/48715e31048aabd8ad2c27d5d647aefcb20a6a53) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Derive `xml:space="preserve"` from the text instead of storing it. `TextContent.preserveSpace` is gone: whether `<w:t>` needs the attribute is a pure function of its text, and a stored copy of a derived fact only drifts — the editor lost it whenever two adjacent runs merged, because ProseMirror has nowhere to carry it. Both serializers now call the same `requiresXmlSpacePreserve`, which `@stll/docx-core` exports.

- [#872](https://github.com/stella/folio/pull/872) [`84c1650`](https://github.com/stella/folio/commit/84c1650475771bc1eeb18ee64ad5605e50089469) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A drawing with no picture relationship keeps the markup it arrived with. A `w:drawing` whose graphic is a chart or an OLE frame, or which carries no `a:graphic` at all, has no `a:blip` and so no relationship id; `Image.rId` is now absent in that case rather than an empty string, and a save writes the anchor back as authored instead of rebuilding it into a picture bound to whichever relationship the part happens to list first. Relationship ids resolve through one typed resolver that distinguishes a resolved id from an absent and a dangling one, and an image reference that names a non-image relationship no longer resolves to that part.

- [#869](https://github.com/stella/folio/pull/869) [`0620ed7`](https://github.com/stella/folio/commit/0620ed7f95babbcc65f2c3a6746da8c7772ed500) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `ExhaustiveFields<Source, Classified>`, the compile-time gate that turns a model field added without a decision into a build failure, now has one owner and is exported from `@stll/docx-core/model`. The paragraph, text, and border serializers each carried a verbatim copy.

- [#868](https://github.com/stella/folio/pull/868) [`170e3ca`](https://github.com/stella/folio/commit/170e3cad6254f7c372c9dff131584df70d3f85be) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `Document.parseWarnings` reports what the parse boundary normalised, as data: a stable `code` from `PARSE_WARNING_CODES`, the part it happened in, the best position that part can name, and the value folio declined to read. `Document.warnings` is unchanged in shape and is now rendered from that list by one formatter, so the prose and the data cannot disagree. Normalisations that were silent now report: a `w:type` outside `ST_HdrFtr`, a repeated footnote or endnote id, a value outside `ST_OnOff` in either shape, a border with no `w:val`, a `w:comment` with no readable `w:id` (previously read as id 0, which manufactured a duplicate), and a hyperlink naming a relationship its part never defined. Retained warnings are capped per code, with the remainder counted.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every `w:tblLook` the author wrote. `CT_TblLook` states which conditional formats a table takes from its table style twice: as the legacy `w:val` bitmask, which `TableLook` had no field for, and as six `ST_OnOff` attributes, which the serializer wrote only when true. A rebuild therefore turned `w:val="04A0" w:firstRow="1" w:lastRow="0" w:noHBand="0"` into `w:firstRow="1"`, and the two are different documents: an absent flag falls back to `w:val`'s bit, an explicit `0` overrides it. `TableLook` gains `val` and each flag is now tri-state. Two readers also disagreed about precedence — the table one OR-ed `w:val`'s bits over an explicit `0`, so a table that switched its header row off got one anyway; `styleParser` now calls the table parser, and `resolveTableLook` is the single place a flag resolves to an answer.

### Patch Changes

- [#871](https://github.com/stella/folio/pull/871) [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the range markers that stand between two blocks. `w:body`, `w:tc`, a header and an SDT's content all admit `w:permStart`, `w:customXml*Range*` and a comment or move range beside their paragraphs, and every block container dropped them on save: a protected range lost its `w:permStart` and the saved file came back unprotected. The markers are now captured verbatim and replayed where they stood, the way an SDT's sibling markers already were. They do not yet survive the editor round trip, which needs a zero-width node rather than a block attribute.

- [#867](https://github.com/stella/folio/pull/867) [`2538df2`](https://github.com/stella/folio/commit/2538df2dc8c6957ca520c9828f776e3bfb216f41) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `nil` and `none` distinct wherever a `CT_Border` is read or written. They are two members of `ST_Border`, not synonyms, and the build-from-scratch serializer rewrote `none` as `nil`. The four `parseBorderSpec` copies also collapse into one reader, so a border element with no `w:val`, an explicit `w:shadow="0"` and the page-border art relationship ids are now read the same way on the paragraph, style, table, cell and page tiers. The light grid a generated table gets when it declares no borders is unchanged, but is now named as the authoring default it is.

- [#866](https://github.com/stella/folio/pull/866) [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compiled legal documents carry outline levels on `ClauseHeading1`–`ClauseHeading3`, so the clause hierarchy reaches Word's navigation pane, a `TOC \u` field and folio's own outline. Without them a compiled agreement had no outline at all.

- [#866](https://github.com/stella/folio/pull/866) [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Markdown headings compile to paragraphs carrying `w:outlineLvl` as well as a `HeadingN` style id, so the result is classified as a heading by outline level even when merged into a document whose own heading styles are localized and that id resolves to nothing.

- [#871](https://github.com/stella/folio/pull/871) [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every attribute a range marker arrived with. `w:moveFromRangeStart` and `w:moveToRangeStart` lost `w:author` and `w:date`, which their schema type requires, so a saved document was markup Word repaired; `w:displacedByCustomXml` was lost on every bookmark, comment range and move range. The markers now model the schema's own `CT_MarkupRange` / `CT_Bookmark` / `CT_MoveBookmark` chain and share one reader and one writer.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open a `w:sym` that names only one of its two optional attributes. `CT_Sym` declares `w:font` and `w:char` optional and Word renders a `<w:sym w:char="F0B7"/>` by falling back to the run's font; folio refused the document twice over, once in the model validator and once in the ProseMirror projection, so a file Word opens did not open at all. Both checks now accept an absent attribute and still reject a malformed character, and the serializer writes an absent attribute back as absent instead of inventing `w:font=""`.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `w:tblGridChange` and `w:numberingChange` when the container they sit in is rebuilt. Both are tracked-change records — the grid a reviewer replaced when resizing a column, and the numbering a reviewer replaced when changing a list — and nothing in the editable model derives either, so a save that rebuilt the container dropped the revision and the document then read as though the change had always been there. They now travel as their own capture slots (`TableFormatting.gridChangeXml`, `ParagraphFormatting.numberingChangeXml`) and are written back on both the replay and the rebuild path. A captured `w:pPr` carrying a `w:numberingChange` is no longer refused for replay either; refusing it used to force the rebuild that could not write it.

## 0.21.0

### Minor Changes

- [#848](https://github.com/stella/folio/pull/848) [`66f0734`](https://github.com/stella/folio/commit/66f0734a93dc0dc78e41dacac4a7414dc9176327) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Classify a drawing as `native`, `replayable` or `opaque` through the predicate the run serializer already uses, so a document is no longer opened read-only because a header carries a logo; only `opaque` content blocks editing, and `DocxCompatibility` gains a `drawings` list at `schemaVersion: 2`. Regenerating a picture now round-trips `a:graphicFrameLocks` and `wp:effectExtent`, and a rasterized shape group is marked `previewOnly` so the editor declines to resize it rather than replacing the group with one child picture. Shape drawings Folio cannot model — unmodeled effects and 3-D, `wpg:wgp` groups without a preview, a `w:pict` with no resolvable image, an `mc:AlternateContent` whose every branch declines — are preserved verbatim instead of dropped. Field results are no longer missing from the AI-facing block text, so a paragraph carrying a cross-reference reads as the text Word shows.

## 0.20.2

### Patch Changes

- [#846](https://github.com/stella/folio/pull/846) [`8fc69b7`](https://github.com/stella/folio/commit/8fc69b78abe4db6bd6078fa9b4ea280137196a44) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `listRendering.levelStarts` through a Document → ProseMirror → Document rebuild so custom list starts render without a DOCX round-trip, and add `formattingScope: "allParagraphs"` to block insertions so a multiline `text` can produce several list items. `ListRendering.levelStarts`, `DocumentSettings.mirrorMargins`, and the header/footer verbatim capture fields are now declared on the model types instead of attached through local intersections.

## 0.20.1

### Patch Changes

- [#838](https://github.com/stella/folio/pull/838) [`301b9e3`](https://github.com/stella/folio/commit/301b9e3442776c63c3a1a392ed6bd0bfbb18dd96) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve editable DrawingML WordArt metadata through parsing, editing, and DOCX serialization.

## 0.20.0

### Minor Changes

- [#836](https://github.com/stella/folio/pull/836) [`e86528b`](https://github.com/stella/folio/commit/e86528b2c70aa2fbc30fdef4561dc1823bb52956) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored run properties through document comparison and tracked revision resolution. Retain text in the selected AlternateContent branch during parsing and serialization.

  Preserve direct paragraph indentation and inline tab/break controls through comparison and reviewed views. Insert operations can retain hard breaks with `lineBreakMode: "inline"`.

  Preserve inherited table run formatting when deleting content. Standalone hard page breaks can be inserted as tracked content with `hardPageBreak`.

  Carry concrete list references and import target numbering definitions. Rebind conflicting definitions through tracked paragraph changes so accepting and rejecting retain the corresponding list formatting.

  Preserve complete table formatting, empty field results, and untouched drawing geometry through document saves. Explicit image edits invalidate stale editable drawing captures.

  Reconcile field, picture, and page-break atoms through mapped review positions. Report atom-only changes and verify both accepted and rejected content, importing picture media without overwriting existing resources.

  Track section property changes through accepted and rejected views, and report unsupported section topology explicitly.

  Add Folio-exact section reference history for reversible header/footer selection changes, with explicit Word save compatibility reporting. Import missing character style definitions during comparison.

  Preserve authored complex-field instructions and import embedded header watermarks without overwriting existing media. Remove retired header/footer parts and exclusively referenced media when their selection changes are resolved.

  Preserve terminal-table review boundaries in Folio-exact mode. Align agent insertion schemas with hard-break exclusivity and positive numbering identifiers.

- [#836](https://github.com/stella/folio/pull/836) [`e86528b`](https://github.com/stella/folio/commit/e86528b2c70aa2fbc30fdef4561dc1823bb52956) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inline and display math inside tracked insertion, deletion, and move wrappers through parsing, editing, and serialization.

  Retain proofing exclusions in run formatting and fingerprint editable raw pictures so untouched DrawingML remains exact while model edits invalidate stale captures.

## 0.19.4

### Patch Changes

- [#819](https://github.com/stella/folio/pull/819) [`7385db9`](https://github.com/stella/folio/commit/7385db9807f293c57e1ca5eae072928bc45a93bf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and render DrawingML text-box rotation and flips across editor and PDF output.

## 0.19.3

### Patch Changes

- [#797](https://github.com/stella/folio/pull/797) [`e289686`](https://github.com/stella/folio/commit/e2896863f28be23473628966a77123573bba2e0b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and recompute imported numbering across list structure edits.

## 0.19.2

### Patch Changes

- [#751](https://github.com/stella/folio/pull/751) [`27e2717`](https://github.com/stella/folio/commit/27e2717fea8dbd367198fdc520ed5ab3364d9828) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Apply section line grids to table cells when the document compatibility setting requests it.

## 0.19.1

### Patch Changes

- [#738](https://github.com/stella/folio/pull/738) [`da95224`](https://github.com/stella/folio/commit/da9522486bdfd5e50c5fa935f10ebb3173b06cd6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A save that changed nothing writes a table's `w:tblPr`, `w:tblGrid`, `w:trPr` and `w:tcPr` back as they arrived, rather than rebuilding them from the typed model and dropping the conditional-format flags, the `w:tblGridChange`, and whatever else the model does not cover. The capture is re-parsed and checked against the model before it is used, so a `Document` edited in place is still honoured. `w:tcPr` and `w:tblPr` also stop gaining an inherited value — a border a table style supplied, a margin the table declared — as the cell's or table's own override, and an absent `w:hideMark` stops being written back as an explicit `w:val="off"`.

## 0.19.0

### Minor Changes

- [#714](https://github.com/stella/folio/pull/714) [`3c5e627`](https://github.com/stella/folio/commit/3c5e627559b2cbd9d06e7c6dd7066488076d67b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Blank paragraphs are blocks. The AI-facing snapshot skipped every paragraph
  with no words, so a blank line had no id, nothing could address it, and a
  comparison could neither add nor remove one. It carries them now, under an
  additive `blank-NNNN` id shape that leaves the published `seq-NNNN` numbering
  over the paragraphs that DO carry words exactly where it was, and the reading
  surfaces filter them through one shared helper.

  `FolioAIEditSnapshot.emptyDocumentAnchorId` is gone, which is why this is a
  minor rather than a patch. It named the anchor an empty document had no block
  for; an empty document is one blank paragraph, and that paragraph is now a
  block, so its id is `blocks[0].id` like any other. Read that instead.

  Three operations follow from that:

  - `insertAfterBlock` / `insertBeforeBlock` with `text: ""` insert a blank
    paragraph rather than being refused as empty.
  - `deleteBlock` on a blank removes it, with its paragraph mark tracked, rather
    than doing nothing.
  - `insertTableRow` sizes the new row against the table's own column count, so a
    row a table can hold is no longer refused because a cell in it spans columns.

  Five defects the blanks made visible, each of which was already losing content:

  - An attribute-only edit — a paragraph mark, a list level, a style on a blank
    paragraph — never reached the saved file. The change tracker reads positions
    off each step, and `AttrStep` carries an empty step map, so the selective save
    wrote the paragraph's original XML and the edit was gone from the document
    while the editor still showed it.
  - Deleting a block left its images behind. The deletion range is built from the
    block's text, so an image outside the words kept no mark, and accepting the
    deletion left a paragraph standing around an orphan picture.
  - Zero-width anchors counted as content. A bookmark boundary, a text-box
    anchor, or Word's cached pagination boundary (`w:lastRenderedPageBreak`)
    survived a deletion that took every word, and the emptied paragraph stayed as
    a blank line nobody asked for. They are not content, they never carry a
    revision, and a revision landing on one produced a document the serializer
    refuses to write.
  - Resolving a paragraph's mark could take a section with it. A section's
    properties live on a paragraph mark, so the paragraph is where the section
    ends: removing it merged two sections into one and dropped the removed
    section's page size, margins, and header and footer references. The
    properties now travel to the paragraph the resolution leaves behind, and a
    paragraph with nothing after it to carry them keeps its mark.
  - Appending a paragraph at the end of a container handed its paragraph mark to
    the anchor. That is equivalent only while the two stay adjacent, and a table
    inserted between them by a later operation in the same batch separated them:
    the mark then joined the anchor to the table, which is nothing, and the
    appended paragraph survived a reject that should have closed it. Every
    inserted paragraph carries its own mark, and resolving a mark with nothing
    after it removes the emptied paragraph instead of joining.

- [#721](https://github.com/stella/folio/pull/721) [`681923a`](https://github.com/stella/folio/commit/681923ab277b78acc69f3eeaea0262ec07fcf168) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve bookmark boundaries inside tracked inline changes so accepting and rejecting revisions keeps bookmark ownership intact.

### Patch Changes

- [#727](https://github.com/stella/folio/pull/727) [`046302a`](https://github.com/stella/folio/commit/046302a9b1b0aef5d7a3dad8d3bf9c33e0f8babd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep physical tracked-change revision IDs unique when saving DOCX packages.

## 0.18.0

### Minor Changes

- [#689](https://github.com/stella/folio/pull/689) [`ab10444`](https://github.com/stella/folio/commit/ab104443c77afde1528148b813c6828db5a2f6e2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Parse legal-source drafts as GFM markdown plus `@` directives with `marked`: clause bodies, list items, and table cells keep inline emphasis, links, and code spans; markdown lists and pipe tables outside a directive become real list and table blocks. `compileMarkdownToContent` and `sanitizeExternalUrl` move into `@stll/docx-core`, and `@stll/folio-core`'s `fromMarkdown` now wraps them.

## 0.17.3

### Patch Changes

- [#667](https://github.com/stella/folio/pull/667) [`a7c9d18`](https://github.com/stella/folio/commit/a7c9d185cd1593e6f34aaf4203f14452ed868d7c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve DrawingML vertical alignment and supported legacy text-box anchors through layout and save.

## 0.17.2

### Patch Changes

- [#637](https://github.com/stella/folio/pull/637) [`4582ad7`](https://github.com/stella/folio/commit/4582ad7671c31e757bbee0ae4d829186dd2be1bc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Index DOCX numbering style links for bounded projection lookup.

## 0.17.1

### Patch Changes

- [#624](https://github.com/stella/folio/pull/624) [`8724016`](https://github.com/stella/folio/commit/872401629e10ca7837bdcf59978d303ef7ccbac5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Parse, render, and round-trip adjusted right-brace shapes as editable geometry.

- [#623](https://github.com/stella/folio/pull/623) [`036b534`](https://github.com/stella/folio/commit/036b534d8c10d81ff084fce781498cb9e511b3b5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allow text editing when unsupported drawings are preserved as inert raw OOXML.

## 0.17.0

### Minor Changes

- [#612](https://github.com/stella/folio/pull/612) [`75d34d2`](https://github.com/stella/folio/commit/75d34d291bd31e40f3c3c35fdbeb3c93529c0653) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve Arabic DOCX parity for no-wrap text boxes, boundary-owned section pagination, RTL tabs, and page-number formats.

## 0.16.0

### Minor Changes

- [#607](https://github.com/stella/folio/pull/607) [`069ba36`](https://github.com/stella/folio/commit/069ba36ac26a686d1b5cbbc2ee69c940cdb238f0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve independent complex-script typography on list markers.

## 0.15.2

### Patch Changes

- [#601](https://github.com/stella/folio/pull/601) [`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use one bounded DOCX package validator with structured failure codes.

## 0.15.1

### Patch Changes

- [#587](https://github.com/stella/folio/pull/587) [`692e0ba`](https://github.com/stella/folio/commit/692e0ba01feabf7f852be6ecde053c7574b1a4d9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Upgrade `better-result` to v3.

## 0.15.0

### Minor Changes

- [#577](https://github.com/stella/folio/pull/577) [`9385d46`](https://github.com/stella/folio/commit/9385d464b85effe1d4aaa558ebea8e29cfcb84e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Headless DOCX report generation: `/server` exports typed builders (`heading`, `paragraph`, `run`, `table`, `pageBreak`, `hyperlink`, `bookmark`, `endnote`, `createTableOfContentsField`); external hyperlinks inside headers, footers, footnotes and endnotes now get relationships in their own rels part; `createEmptyDocument` initialises `package.relationships` so in-memory headers and footers materialise; `DocumentSettings.updateFields` round-trips; the Stella style set gains `Heading1`-`Heading6`, `TOCHeading`, `TOC1`-`TOC3`, `EndnoteReference` and `EndnoteText`; complex fields keep `w:dirty`/`w:fldLock` across parse and save.

## 0.14.0

### Minor Changes

- [#569](https://github.com/stella/folio/pull/569) [`a38d902`](https://github.com/stella/folio/commit/a38d9025773e1bdba6bfbb4ffcab3dea1a943d6e) Thanks [@berticeek](https://github.com/berticeek)! - Project effective paragraph alignment (`w:jc` from direct, style chain, and docDefaults) as a seventh paragraph tuple slot; `DOCX_PROJECTION_SCHEMA_VERSION` moves from 4 to 5. `start`/`end` and other unsupported `w:jc` values project as absent; table-style and numbering-level alignment are not consulted, so a paragraph aligned only by those tiers projects the lower-tier style or docDefaults value.

## 0.13.0

### Minor Changes

- [#537](https://github.com/stella/folio/pull/537) [`886b6f6`](https://github.com/stella/folio/commit/886b6f6cd0f2a407c872c90c2ef294192ea3bc0c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project effective text formatting, visible Office Math text, font-bound symbols, and bookmarks at document or table boundaries from DOCX packages. Style and numbering parts follow the main document's OPC relationships, and formatting projections report when unsupported package style layers prevent their spans from being authoritative.

## 0.12.0

### Minor Changes

- [#534](https://github.com/stella/folio/pull/534) [`2ef7445`](https://github.com/stella/folio/commit/2ef74452c34153a09b1af7496f27d8abd2074efc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Flatten attributed review facts at the WebAssembly boundary to reduce projection allocation and retained heap.

- [#536](https://github.com/stella/folio/pull/536) [`51cbe7d`](https://github.com/stella/folio/commit/51cbe7d708fe44786a7fe8165c42225d2e35b93e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project effective paragraph indentation from bounded OOXML numbering definitions.

## 0.11.0

### Minor Changes

- [#529](https://github.com/stella/folio/pull/529) [`ae98003`](https://github.com/stella/folio/commit/ae98003cf69b31ce86b44b5be0ff30834ee71455) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Project direct superscript formatting from DOCX runs.

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
