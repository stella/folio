# @stll/folio-core

## 0.47.2

### Patch Changes

- [#943](https://github.com/stella/folio/pull/943) [`371ff8f`](https://github.com/stella/folio/commit/371ff8f13f0c04381428c2093529376affcc2cd8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored hyperlink and note-reference metadata across editor saves.

- [#951](https://github.com/stella/folio/pull/951) [`6f8e61e`](https://github.com/stella/folio/commit/6f8e61e937960a521d318a366465dec5e79eec64) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop counting markup between runs as a run-formatting carrier.

  A preserved capture is a `w:r` child or a `w:p` child, and only the first serializes inside a run and owns a `w:rPr`. Both arrive as the same editor node, and the map from node type to run-formatting carrier could only name the type, so the paragraph-level capture — `w:proofErr` above all, which Word writes between the runs of any sentence its grammar checker flags — was classified as a run. Formatting marks put on it were dropped on the way back out, and a comparison was asked to line the base document's proofing annotations up with the revised document's own: no redline can, so the round-trip check refused redlines whose content was right, reporting `inline-formatting`. The classification now reads the level the capture came from, which is the level the save path already branches on.

- [#944](https://github.com/stella/folio/pull/944) [`55229a3`](https://github.com/stella/folio/commit/55229a38e4906fb471749d3f5420d7d0fcc9a8d0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored numbering-format metadata for lists and note numbering.

- [#945](https://github.com/stella/folio/pull/945) [`dade6b3`](https://github.com/stella/folio/commit/dade6b363053471cb26852b1b49fc93e5e09360b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stabilize the public declaration for table-property justification values.

- [#945](https://github.com/stella/folio/pull/945) [`dade6b3`](https://github.com/stella/folio/commit/dade6b363053471cb26852b1b49fc93e5e09360b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored table-cell identifiers through editor saves.

- [#949](https://github.com/stella/folio/pull/949) [`f7081b0`](https://github.com/stella/folio/commit/f7081b0bfecd5b94a2bd7df63933bfe4c4ff9b01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored carriage-return elements through editor saves.
- Updated dependencies [[`371ff8f`](https://github.com/stella/folio/commit/371ff8f13f0c04381428c2093529376affcc2cd8), [`55229a3`](https://github.com/stella/folio/commit/55229a38e4906fb471749d3f5420d7d0fcc9a8d0), [`dade6b3`](https://github.com/stella/folio/commit/dade6b363053471cb26852b1b49fc93e5e09360b), [`f7081b0`](https://github.com/stella/folio/commit/f7081b0bfecd5b94a2bd7df63933bfe4c4ff9b01)]:
  - @stll/docx-core@0.25.1

## 0.47.1

### Patch Changes

- [#946](https://github.com/stella/folio/pull/946) [`8b98daa`](https://github.com/stella/folio/commit/8b98daa0cda9e6d051bccf4079761b2490940a08) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop a comparison from recording a run-property change that changes nothing.

  Replacing text clears a highlight or `w:shd` under it, because text typed over a highlighted placeholder is new text and the marker that said "fill this in" should not survive into the finished document. A comparison is not authoring: it holds the revised document's own run properties and writes them itself. Clearing them first recorded a `w:rPrChange` that the provenance pass then took straight back, so every carrier of an edited paragraph in a highlighted cell reached the reader as a revision whose before and after were identical, and the run the replacement deleted kept a claim that its background had gone.

  `replacementBackground` names the two behaviours on the apply path, defaulting to `clear`; `compareDocx` asks for `keep`.

- [#946](https://github.com/stella/folio/pull/946) [`8b98daa`](https://github.com/stella/folio/commit/8b98daa0cda9e6d051bccf4079761b2490940a08) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align a cell's paragraphs on each side of its nested table separately.

  A cell's paragraphs were aligned as one sequence, so a cell that lost paragraphs could pair a surviving one with a paragraph on the far side of its nested table. Nothing moves a block across a table, and the cell's last paragraph cannot be deleted because its mark has nothing to join, so the redline kept a paragraph the revised document does not have and the comparison refused its own round trip with `container`. The paragraphs between a cell's nested tables are now aligned run by run, the way rows and cells already are.

  A cell whose last block is a table states a cell no consumer renders as written: `CT_Tc` ends in a paragraph, and every consumer reads the implied empty one there. Producers leave that shape behind when they rewrite a cell and delete the closing paragraph along with the rest. Parsing it as written asked the comparison to delete a paragraph mark that has to stay, so the parser now reads the implied paragraph as the fact it is, the way it already supplies one for a cell that states no block at all.

## 0.47.0

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

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write `w:tcW` from the width a cell states, not from the width the table resolved for it.

  `TableCellAttrs.width` is the width the cell renders at: when the cell declares no `w:tcW`, the table resolves one from its grid and puts it there. The way back wrote that attr into `w:tcPr` unconditionally, so opening a document and saving it again gave every cell a preferred width its author never wrote.

  `TableCellAttrs._authoredWidth` records what the cell itself states, as `_resolvedBorders` and `_resolvedMargins` already do for the border and the margin, and the save leg writes `w:tcW` from it alone. A command that moves a cell's width states one: `mergeTableCellAttrs` derives the record for every command that patches a cell, so a column resize writes exactly the cells it moved and a merge sums the widths its source cells stated rather than the widths the grid gave them.

  The attr-schema version moves to 4. A version-3 snapshot states no `_authoredWidth`, and `_originalFormatting.width` is its record of which cells wrote a `w:tcW`, so the step backfills from it at the width the cell currently holds.

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a comment or tracked move anchored on a point, at the point it was anchored at, through the editor.

  Word writes a comment left at an insertion point as `w:commentRangeStart` immediately followed by `w:commentRangeEnd`, then the `w:commentReference` run; a tracked move whose range covers nothing is spelled the same way. The editor projects a comment as a `comment` mark over the content its range covers, and a move range as a marker placed around the move's wrappers, so a range with nothing between its two markers had no carrier: the comment still listed, because its reference is an atom of its own, but the anchor position was gone and the save wrote a reference with no range. An empty move range came back around the whole paragraph, which is where the marker carrier puts one whose wrappers it cannot find.

  The editor gains a zero-width `rangeAnchor` node holding the two markers, so the pair comes back adjacent at the position it was authored at, inside the wrapper it was authored inside. The pair is one node rather than two boundary nodes: two would leave a position between them for a caret, and typing there would widen a range the reviewer drew as a point. Deleting the anchor removes the comment, exactly as deleting its reference does. A range with any content in it, including one whose content is only a bookmark boundary or a capture, is untouched and still travels as the mark.

  The node is additive: editor state written before it holds none and needs no migration.

  Sixty container-census pairs — the six range markers across the ten inline containers that declare them — are now `modelled`, thirty-six of them moving from `dropped (editorProjection)` here.

- [#938](https://github.com/stella/folio/pull/938) [`9d1e897`](https://github.com/stella/folio/commit/9d1e8973631997ee478ca6ad0fedda4cc1134246) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read `w:sdtPr` through the shared child dispatcher, so a content control keeps every property its author wrote. `SdtProperties.preserved` holds the children folio does not model at their `CT_SdtPr` ordinal, and one writer serialises block, inline, row and cell controls from the model rather than replaying the source's bytes. `SdtProperties.rawPropertiesXml` is gone; `SdtProperties.lock` no longer reports `unlocked` for a value the reader refuses.

- [#927](https://github.com/stella/folio/pull/927) [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Draw a `PreviewDescriptor` instead of rasterising one. A diagram's shapes become filled `rect` primitives in the image's box, so the DOM backend paints the absolutely positioned divs it paints every rectangle as and the PDF exporter emits vector operators rather than embedding a bitmap of a vector drawing. `ImageTable` no longer interns a descriptor, and `MAX_BUILD_PREVIEW_PIXELS` is gone with the rasters it bounded. A shape now keeps the colour the drawing authored, including a channel of zero, which the raster read as the backdrop's.

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

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - One `ListState`, resolved from the numbering definitions

  The toolbar's list state is a union in `@stll/folio-core/prosemirror`, read by
  both adapters, in place of three declarations that never imported one another
  and had already drifted. It replaces the `isInList` flag, which restated
  `type !== "none"`, and the `numId` that meant nothing on the empty arm.

  Which kind of list a paragraph is in now comes from its level's `w:numFmt`
  rather than from its numbering id. `numId === 1` meant bullets only in a
  document Folio had created itself, so an imported bulleted list read as
  numbered in every toolbar. The list commands keep minting their own instances
  and say which kind they are creating, for a document that defines no numbering
  yet.

  `SelectionState` carries the resolved `listState`, and `SelectionContext`
  replaces `inList` / `listType` / `listLevel` with it.

- [#931](https://github.com/stella/folio/pull/931) [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate `NumberFormat` from `ST_NumberFormat`. It omitted `bahtText`,
  `dollarText` and `custom`, and carried three `decimalZero{3,4,5}` members the
  format does not declare: the parser minted them from a custom format's pad
  width and the serializer wrote them back as a `w:val` no consumer can read.

  A custom format is now held as `custom` plus the `@w:format` it counts by
  (`ListLevel.numFmtFormat`), and written back as both. The three synthetic
  values move to `CounterFormat`, the render vocabulary `ListRendering.numFmt`
  and the editor's list attributes carry, which is never serialized.

- [#936](https://github.com/stella/folio/pull/936) [`3fde3f4`](https://github.com/stella/folio/commit/3fde3f47fcdf038dd4a787a65fb46c4342752da1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a numbering level's `w:pPr` with the reader the other three owners of the set already share.

  A level kept a private copy that took an indent and a tab list and let the other thirty-one children `CT_PPrGeneral` declares fall off the end of the walk, while the level's writer was already the shared one: a `w:pStyle`, a `w:jc`, a `w:spacing` or a `w:keepNext` on a level was read as nothing and written back as nothing.

  The reader could not be shared before because `paragraphParser` and `numberingParser` import each other, so it now lives in `docx/paragraphProperties.ts`, which imports neither and which both import.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry `w:numPr` as a union through the editor attr

  The persisted `numPr` and `numPrFromStyle` paragraph attrs hold
  `ParagraphNumberingOverride`, the shape the model already carries, minted by
  one codec. `toProseDoc` and `fromProseDoc` stop converting between two slot
  pairs and a union at the editor boundary, and a recorded `w:pPrChange` stores
  the same union with `null` still meaning "carried no numbering".

  `FOLIO_YJS_ATTR_SCHEMA_VERSION` is 5. A stored version-4 snapshot is carried
  forward by `migrateFolioYjsSnapshot` and by every load path: `w:numId` 0
  becomes the cancellation, an `w:ilvl` without an id becomes the level-only
  arm, the pair becomes a reference, and an element that stated neither slot
  stays absent.

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

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `w:trPr/w:hidden` is tri-state: absent, an explicit on, an explicit off.

  `tableRow`'s `hidden` attr defaulted to `false`, so the editor could not tell a row that authored `<w:hidden w:val="0"/>` from one that authored nothing, and the parser had to let an explicit off travel as captured bytes rather than model it. The default is `null` now, as `heightRule`'s already is, and the parser reads all three states into `TableRowFormatting.hidden`.

  The attr-schema version moves to 5: a version-4 snapshot's `false` never meant an explicit off, so keeping it as one would start writing an element the document never carried, and the step drops it.

- [#940](https://github.com/stella/folio/pull/940) [`f1a4d2d`](https://github.com/stella/folio/commit/f1a4d2dd55fc83bc3872253fd4b00a017785ec85) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry a run's own record through the editor on one mark.

  An authored `w:r` had no record on the editor side, so a round trip rebuilt it from formatting alone: its attribute remainder — the `w:rsid*` family Word writes on nearly every run — and the `w:rPr` children no reader took a value from were both dropped, and two adjacent runs the formatting could not tell apart were written back as one.

  `runIdentity` is that record. It is one mark rather than three, because all three payloads are facts about one element and the save leg asks that element a single question: where does this run begin and end. It is in the key adjacent leaves are grouped by, so a change of identity is a run boundary; it is minted only when a run holds a page break, a remainder or a sink, so a fully modelled document pays nothing. It replaces `pageBreakRunOwner`, whose job was the same one level narrower.

  Two rules follow from what an rsid means. It names an editing session registered in the package's own `settings.xml`, which folio neither writes nor merges, so folio writes no rsid it did not read: text typed inside an authored run becomes a run of its own that states none, and the paragraph's `w:rsidRDefault` answers for it. And a pasted span carries the id alone, so a copy keys differently from its source and becomes its own run. Splitting a run, by contrast, manufactures nothing, so both halves keep what the run was authored with.

  Collaboration snapshots move to attr schema 4. The rename is a value rewrite rather than an additive change, because mark attrs live in the shared text's delta under the mark's own name and an unknown name costs the text beneath it, not just the mark. The step rewrites the delta attribute before anything reads the fragment, and the marker turns an older build meeting a newer snapshot into a refusal. Payloads are not backfilled: a snapshot has no access to the package it was seeded from, so an existing room keeps today's behaviour until it is reseeded.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every `w:rPr` child folio does not model, where it stood, for a run, a paragraph mark, a style and a tracked property change alike.

  A run property set was walked by a reader per property with no branch for the rest. `w:bdr`, `w:fitText`, `w:eastAsianLayout`, `w:snapToGrid`, `w:webHidden`, `w:specVanish` and `w:oMath` had no model at all; `w:rFonts`, `w:u`, `w:lang`, `w:w` and `w:sz` were read and dropped whenever the reader took no typed value from them; and the paragraph mark's whole `w:rPrChange`, along with everything inside it, went with them. A save that rewrote the element — which is every save after an edit — lost all of it.

  `EG_RPrBase` now goes through the shared child dispatcher, with a handler map the compiler makes total over the children the schema declares. A handler answers with what it took, so a property the reader turned into no typed value keeps its bytes: a name-keyed map can state the names folio has never heard of, not the values a reader refuses. `TextFormatting` gains `preserved`, and because it is a sequence the sink records each capture's schema ordinal rather than a count of modelled siblings.

  The four owners of a run property set — a run, the paragraph mark inside `w:pPr`, and the snapshot inside either one's `w:rPrChange` — share that map and differ only in which children a sibling record has already claimed, which the call site names. One writer serves all of them plus a style and a numbering level, and it orders its children from the generated declared-child list rather than from the order of its own statements: folio wrote `w:vanish` before `w:noProof` while the schema declares the reverse, which a validating consumer refuses.

  Captured bytes belong to the element that was parsed and to no other, so style resolution and formatting merges drop them rather than inheriting them onto every run below.

- [#939](https://github.com/stella/folio/pull/939) [`63c18a8`](https://github.com/stella/folio/commit/63c18a81c3bad202d054f2e7cf32367df9c1c44c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Lay out a `nextColumn` section break. `ST_SectionMark` has five members and the layout modelled four, so `w:type="nextColumn"` reached the paginator through a cast as a value its switch did not handle: the section fell through to the switch's silent default, started no section at all (its page numbering and header references never took effect) and resumed below the outgoing content instead of in the next column.

  `SectionBreakBlock["type"]` is now `SectionStart`, the model's `ST_SectionMark`, so the bridge hands the parsed value through and a member the layout does not handle is a compile error rather than a cast. Per §17.18.77 a `nextColumn` section begins in the next column of the region it shares with the outgoing section; when there is no such column, because the section is single-column or because the incoming one redefines the column geometry, it begins in place like `continuous`, and when the region's last column is already in use it opens the next page. The measure pass mirrors the same decision, so the width and column a block is measured against match the one it is painted in.

  `normalizeSectionBreakType` still reads an absent `w:type` as `nextPage` (§17.6.22) but no longer passes a value outside the enumeration through as though it were one: that is a bug in whatever produced it, and it now panics instead of laying out as something else. `SECTION_BREAK_TYPES` is total over the enumeration, checked against the committed schema graph, and the Yjs v3 migration therefore keeps a `nextColumn` section rather than dropping it. Which members the insert commands offer is unchanged.

- [#925](https://github.com/stella/folio/pull/925) [`da1fc6a`](https://github.com/stella/folio/commit/da1fc6a4c8d7a48705187d164d8e8180cefeff74) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a `w:lnNumType` or `w:pgBorders` that states nothing, on the live section and inside a `w:sectPrChange`.

  Both types declare every attribute and every child optional, so the element alone is a legal document: it says the section is line numbered, or bordered, with the defaults. The parser recorded it and the two serializers wrote nothing for it, which `serializeDocGrid` in the same file had already decided the other way — an attribute-less element is a document folio must write back, and the record exists only because the parser read one.

  The loss showed up differently in the two places the element can sit. On a live section the record then held a field the serializer did not write, so the package fidelity guard refused the save outright. Inside a `w:sectPrChange` the refusal is swallowed and the change is written with an empty `<w:sectPr/>`, so the snapshot a reviewer would restore came back blank.

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

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry a stored `outlineLevel` node attr forward to the union shape.

  `FOLIO_YJS_ATTR_SCHEMA_VERSION` moves to 4, and the version-3 step rewrites every paragraph's `outlineLevel` from the `w:outlineLvl w:val` number it stored into `OutlineLevel`: 0..8 become the heading arm, 9 becomes the body-text arm, and a value the format never defined is dropped, matching the parse boundary.

  The step cannot be skipped and read lazily. ProseMirror copies a stored attr into the node without validating it, so an unmigrated snapshot would reach the strict validator as a bare number; the version gate fires first, on both the editor and the materialization load paths, which is what turns a silent misread into a rewrite.

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

- [#938](https://github.com/stella/folio/pull/938) [`9d1e897`](https://github.com/stella/folio/commit/9d1e8973631997ee478ca6ad0fedda4cc1134246) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry a table's and a row's verbatim sink through the editor, not only through a save.

  `Table.preserved` and `TableRow.preserved` hold the markup a `w:tbl` carried beside its rows and a `w:tr` beside its cells — a bookmark or permission boundary, a proofing error, a custom-XML revision range — with the count that places each back between the same two siblings. The save leg wrote them; `toProseDoc` had nowhere to put them, so opening a document and saving it dropped them. A `w:bookmarkEnd` written after a table's last row is the case that shows: losing it leaves the `w:bookmarkStart` in a cell with no end.

  `TableAttrs` and `TableRowAttrs` gain `_preserved`, carried by reference the way `_preservedAttributes` is, so a table or row the editor created has no sink and a copy does not inherit one. `preservedSinkCarriers.ts` asks the question once per model record that declares a sink, over a union derived from the model rather than listed, so the next sink cannot reach the editor without an answer.

- [#932](https://github.com/stella/folio/pull/932) [`166d3f0`](https://github.com/stella/folio/commit/166d3f0868393dc74c0609ac4a94cba0e579743b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Pair `ST_TextDirection`'s two spellings of each flow by ECMA-376 Part 4
  §14.11.7. Folio paired them by the letters in the token, so every one of the
  six Strict spellings rendered as something other than its Transitional twin:
  `tb` is the horizontal flow and turned a quarter clockwise, `rl` and `lr` are
  vertical flows and painted flat. Rendering is now decided per flow, and the
  section's own text direction is narrowed against the enumeration rather than a
  second hand-written copy of it.

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

- [#913](https://github.com/stella/folio/pull/913) [`94853aa`](https://github.com/stella/folio/commit/94853aa3a42f02fa56079f571ace5368fba3efea) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep an edit local in a package whose producer wrote no `w14:paraId`. The
  selective save keyed paragraph identity on the id alone, and folio mints one for
  every paragraph that arrives without one, so in a LibreOffice, Google Docs,
  python-docx or docx4j document every paragraph looked new: the splice was
  declined, the whole of `word/document.xml` was rebuilt, and the minted ids were
  written to disk. `resolveParagraphIdentities` now decides each paragraph's
  identity once — `authored` when the source part writes the id, `minted` when it
  does not — and the patcher consumes that union exhaustively, addressing an
  authored paragraph by id and a minted one by its ordinal, which the part's
  remaining authored ids prove. A save no longer stamps a minted id into the
  package; `ensureParaIds` remains the pass that gives a package ids, at ingest,
  when a host asks for it.

- [#917](https://github.com/stella/folio/pull/917) [`8f9a01b`](https://github.com/stella/folio/commit/8f9a01b2b849e743a3dbac019e7a1aecbf9c2379) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `wp:anchor`'s `simplePos`, `relativeHeight`, `locked` and `hidden`, the `wp:simplePos` offsets, and both `wp:docPr` links whole when an edited drawing is rebuilt from the model.

- [#940](https://github.com/stella/folio/pull/940) [`f1a4d2d`](https://github.com/stella/folio/commit/f1a4d2dd55fc83bc3872253fd4b00a017785ec85) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write an explicit off for every `CT_OnOff` element. `serializeOnOffElement` is the
  one writer: absent writes nothing, an on writes the bare element, and an off
  writes `w:val="0"`, which is what cancels an inherited on. The row, cell, table,
  control and paragraph-mark readers keep the three states apart as well, and a
  control that states nothing keeps stating nothing through the editor.

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the revision that covers a whole inline content control.

  `w:ins > w:sdt` reached the editor as a control whose leaves carried nothing. The control is an `inline*` node rather than an inline atom, so it is not a run carrier and the revision mark had nowhere to land; the save leg then wrote the control back beside the revision that had held it, and text a reviewer had inserted was no longer inserted.

  The revision now rides the leaves the control holds, and a revision that covers all of them is written back around the control. Accepting or rejecting it is an operation over the control itself: rejecting an inserted control removes it rather than leaving an empty one standing where it was, on the editor-command path and on the headless one alike. A revision over part of the content has no such form and stays where the editor holds it, per child.

  `w:ins > w:sdt` and `w:sdt > w:ins` reach the editor as the same marks on the same leaves, so a span rebuilt from the editor comes back revision-outermost, as it already does for a hyperlink and for a transparent wrapper. A paragraph nobody edited keeps its authored order, because selective save replays its bytes.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Model `w:outlineLvl` as a union, so the reserved body-text value cannot be read as a tenth heading level.

  `ParagraphFormatting.outlineLevel` was `number`, which put ECMA-376 17.3.1.20's rule ("9 specifically indicates that there is no outline level applied to this paragraph") in every consumer's hands. `OutlineLevel` is now `{ kind: "bodyText" } | { kind: "heading"; level: 0..8 }`, with `level` a union of nine literal types: the sentinel has no representation as a heading, an out-of-range value has none at all, and an absent field still means "states none, inherits one".

  One reader owns the parse boundary (`outlineLevelFromStatedValue`) and one writer owns the emit (`outlineLevelStatedValue`). The paragraph parser, the style parser, the style cascade, the display-list outline, the layout bridge, the ProseMirror attr and its validator, markdown, the style sets and the legal-source compiler all move to the union; `isHeadingOutlineLevel` and the bare `BODY_TEXT_OUTLINE_LEVEL = 9` are gone, replaced by `headingLevelOf` and the body-text arm.

  A `w:outlineLvl` outside 0..9 is now dropped at the parse boundary rather than carried through the model, which is what the Rust projection kernel already did. The container-survival census records the one value that stops surviving a rebuild.

- [#942](https://github.com/stella/folio/pull/942) [`780ecd0`](https://github.com/stella/folio/commit/780ecd0d8d73353bc0d930d97e525d4b516187ac) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `fromProseDoc` takes the reuse it is asked for. A third argument,
  `{ reuse?: ProjectionReuse }`, names whether a record the editor did not change
  may come back from the base document by reference. `"none"` is the default and
  is what the function has always done: every record is rebuilt out of
  ProseMirror. `"matched"` is the merge against a matched base record, and it
  panics until it is implemented rather than falling back to a rebuild, so a
  caller cannot believe it asked for a merge and get today's behaviour.

  The option ships before the merge because the measurement has to. The corpus
  gate's new `editor-projection` invariant is `editor-round-trip` with
  `{ reuse: "none" }` forced, the way `reserialize` strips the capture slots so
  the serializers must run. Once reuse lands, `editor-round-trip` measures the
  merge and `editor-projection` measures the projection; until then the two are
  the same measurement, which is the point of adding the second one first.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Model `w:numPr` as a union

  `ParagraphFormatting.numPr` and `numPrFromStyle` carry
  `ParagraphNumberingOverride` instead of two optional slots, so the reserved
  `w:numId w:val="0"` has no representation past the parse boundary and a level
  stated without an id is a named arm rather than a half-filled pair. The
  cascade fold `mergeParagraphNumbering` replaces the object spreads that used
  to restate ECMA-376 17.3.1.19 at each tier.

- [#939](https://github.com/stella/folio/pull/939) [`63c18a8`](https://github.com/stella/folio/commit/63c18a81c3bad202d054f2e7cf32367df9c1c44c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry a section break on one attr.

  Two paragraph attrs carried one state: `_sectionProperties`, the parsed record the save leg preferred, and `sectionBreakType`, which the editor commands wrote. `removeSectionBreak` cleared the type and left the record, so a parsed break it claimed to remove was still saved; `insertSectionBreak` could not retype a parsed break without the save leg minting a fresh record and dropping that section's page size, margins, columns and header references; and a paragraph holding the type alone had no record for a split to be read against, so both halves minted their own and the document gained a section.

  `_sectionProperties` is now the whole state. The break type is a field of the record (`w:type`, ECMA-376 Part 1 §17.6.22), derived through `sectionBreakTypeOf` for the layout bridge, compare, the change tracker, the DOM and the toolbar commands, and `ParagraphAttrs.sectionBreakType` is gone from the schema. A break the editor inserts mints one record and shares it by reference, exactly as a parsed one is shared, so the save leg's rule (among the paragraphs holding one record, the last in document order writes it) covers both.

  Backspace at the start of the paragraph _after_ a break now deletes the break, as Word does: §17.6.18 puts the section's properties on the mark the join consumes, and the paragraphs it governed fall to the following section, whose `w:sectPr` governs them from then on.

  Collaboration snapshots carry an attr-schema version, bumped to 9. An older snapshot that states `sectionBreakType` and no record has the record minted for it on load; without the step ProseMirror would drop the attr its schema no longer declares and the save would write one `w:sectPr` fewer.

- [#918](https://github.com/stella/folio/pull/918) [`3b984e5`](https://github.com/stella/folio/commit/3b984e5759f40fe5af1c658baa9c163078b7db69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read and write `wp:anchor`'s own attributes from one `DrawingAnchor` record shared by pictures, shapes and text boxes, and carry it through the editor.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give `w:numPr` one union, one cascade fold and one reader.

  `ParagraphNumberingOverride` is what a tier states — `none` (the reserved `w:numId 0`), `reference` (an id and an optional level) or `levelOnly` — and `ResolvedParagraphNumbering` is what the cascade leaves. Three arms, not two: `w:numId` and `w:ilvl` inherit independently (ECMA-376 17.3.1.19), so a tier that states only the level keeps the id it inherits, and that shape is what Word writes whenever a styled list paragraph is demoted.

  `mergeParagraphNumbering` is that inheritance written once. It replaces the paragraph parser's object spread, and it is closed under itself and associative over the three cascade tiers, so a third tier needs no special case.

  `paragraphNumberingFromSlots` is the one mapping from the element's two slots onto an arm, and `readParagraphNumbering` reads the element. Both are exported from `@stll/folio-core/docx` alongside `NO_NUMBERING_NUM_ID` and `isNumberingReference`, which now live in `@stll/docx-core` where the model does. Three duplicate spellings of the reserved id are retired: the hand-inlined copy in `docx-core`'s validator, the bare literal in the operation reader, and the relational form in the AI snapshot, which was the one spelling that read a malformed package's negative id as "not numbered" while every other spelling read it as a dangling reference.

- [#925](https://github.com/stella/folio/pull/925) [`da1fc6a`](https://github.com/stella/folio/commit/da1fc6a4c8d7a48705187d164d8e8180cefeff74) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write a `wp:effectExtent` back on the element that authored it, so a wrap child's own reservation survives a rebuild.

  `CT_Inline` and `CT_Anchor` declare a `wp:effectExtent`, and so do `CT_WrapSquare` and `CT_WrapTopBottom`. They are two values: the drawing's is the object's own effect reservation, the wrap child's is the reservation the text flow is computed against. folio read only the drawing's, into `Image.padding`, and wrote it back there, so a wrap child's own reservation round-tripped an untouched document on the strength of its captured bytes and was gone the moment anything forced the serializer.

  `ImageWrap` gains `effectExtentSlots`, the `distanceSlots` shape one element over: `drawing` and `wrapChild`, each holding the element's four sides. The value in force stays where its consumers read it, on `Image.padding`. `resolveEffectExtents` decides the rebuild the way `resolveWrapDistances` decides the insets — each reservation goes back on the element that stated it while the drawing's is unmoved, and once an editor has resized it the rebuild states the value in force on the drawing alone rather than keeping a wrap reservation computed against a shape that is no longer there. The slots ride through the editor as `wrapEffectExtentSlots` on the image, shape and text-box nodes.

  A shape and a text box have never had a reservation of their own — the rebuild wrote `l="0" t="0" r="0" b="0"` on every one of them — and now keep the one they were authored with.

### Patch Changes

- [#942](https://github.com/stella/folio/pull/942) [`780ecd0`](https://github.com/stella/folio/commit/780ecd0d8d73353bc0d930d97e525d4b516187ac) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop the block alignment from reporting a crossing as a rewrite.

  Anchoring is monotone, so two paragraphs that swap places produce two exact-text correspondences that cross and only one can be anchored. The leftover ends were then handed to the positional fallback, which pairs by offset and knows nothing of what it pairs: a heading and an unrelated paragraph came back as a replacement, taking with it the removal and the arrival the move pass exists to pair.

  An exact correspondence the monotone pass had to drop is still evidence, and it is evidence against fusing that block with something else. The fallback now declines a pair when each side's text stands unpaired on the other side, so the relocation is reported as one.

  Leaving the two ends unpaired reaches a shape the terminal-carrier repair did not recognise: the story's last paragraph removed and another written where it stood. The rotation that carries an inserted mark at a container's end turns on the paragraph the run was appended after, so a paragraph that is itself deleted leaves the addition nowhere to go, and the comparison refused its own round trip. That repair now covers it.

- [#918](https://github.com/stella/folio/pull/918) [`3b984e5`](https://github.com/stella/folio/commit/3b984e5759f40fe5af1c658baa9c163078b7db69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve a drawing's `wp:inline` / `wp:anchor`, its wrap element and its position by namespace URI rather than by the `wp` prefix.

- [#917](https://github.com/stella/folio/pull/917) [`8f9a01b`](https://github.com/stella/folio/commit/8f9a01b2b849e743a3dbac019e7a1aecbf9c2379) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a `w:fldChar` whose complex field the paragraph never closes, along with the run and the text around it, and preserve the `w:ffData` a legacy form field hangs on it.

- [#927](https://github.com/stella/folio/pull/927) [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Charge a WordprocessingGroup preview to the package preview budget. The group producer stamped its own data-URL prefix, mime type and filename, so the budget never recognized one and a package retained as many group previews as it happened to contain. The producer now builds its image from the table, and a package past the allowance has the preview dropped, keeping the drawing, its page space and the authored XML the package saves from.

- [#930](https://github.com/stella/folio/pull/930) [`56539e3`](https://github.com/stella/folio/commit/56539e3504cf2ed26e6b2e016bd66507581d5c24) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read `w:themeColor`, `w:themeFill` and `a:schemeClr/@val` through the generated
  enumerations, so `hyperlink`, `followedHyperlink`, `dark1`, `light1`, `dark2`,
  `light2` and `none` survive a save instead of being dropped at parse time.
  `w:shd` keeps its pattern colour's theme reference too.

- [#940](https://github.com/stella/folio/pull/940) [`f1a4d2d`](https://github.com/stella/folio/commit/f1a4d2dd55fc83bc3872253fd4b00a017785ec85) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read `w:hyperlink@w:history` as the three states it has, so an explicit `w:history="0"` survives a save instead of reaching disk as nothing.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Type the paragraph attr patches a style or a list command produces. `paragraphAttrsFromResolvedStyle`, `listAttrsFromResolvedStyle`, `listAttrsFromNumbering` and `listLevelAttrPatch` returned `Record<string, unknown>`, so every attr they wrote was unchecked: a `numPr` taken straight off a model object assigned as readily as one minted by `paragraphNumberingAttr`, which is the crossing the `ParagraphNumberingAttr` brand exists to refuse, and a misspelled attr key was a silent no-op.

  They now return `ParagraphAttrsPatch`, derived from `ParagraphAttrs` rather than hand-listed: every key carries its own attr type or the absent state the node spec's default stores, and an attr added to the spec is writable without a second edit. A compile-time proof beside the producers pins that the model's union, a hand-assembled record, and an undeclared key are all rejected.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Merge two adjacent runs only when their whole record agrees, and keep the record the merge inherits.

  `consolidateRuns` compared typed formatting alone and rebuilt the merged run from `type`, `formatting` and `content`, so a run's attribute remainder and its property set's verbatim sink were dropped at parse time, before anything downstream could see them. Two runs written in different editing sessions became one run in neither, and where the sink differed the survivor's captured bytes were applied to the other run's text as well: a `w:webHidden` on one run hid the text beside it.

  One predicate, `runsMergeable`, now answers the question for every consolidation site: formatting, attribute remainder and `w:rPr` sink must all be equal. The remainder compares as a set, because attribute order in XML says nothing; the sink compares as a sequence, because its order is what puts the markup back between the same modelled siblings. The merged run is the survivor spread whole, so no field can be dropped by a field list that forgot it.

  Both field lists the predicate reads are now total over their model type, which turned up a second field the merge had been crossing: `w:noProof`.

  More of a document's runs survive a parse as a result, by roughly a third on files that carry run-level session ids.

- [#922](https://github.com/stella/folio/pull/922) [`d098792`](https://github.com/stella/folio/commit/d098792ef0634154446a13d80757b7f73b233838) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the words a link or a field nested inside one of its own kind puts on the page.

  `CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`, so either may hold another of itself. folio does not model the nesting — Word writes neither, and the public corpus has four nested links in one package written by a converter — and captured the inner element through the child sink, which knows nothing about a capture beyond its bytes. So the markup survived and the text did not: a link inside a link, and a field inside a field's cached result, reached the editor as an opaque atom showing nothing, and text extraction, markdown and layout skipped the words entirely.

  Both are captured through the element instead, the way `w:customXml` and `w:smartTag` already were, so the capture carries the visible text beside the markup. Position is unchanged: the capture is still a member of the owner's own content union, between the same two children it was read between.

- [#914](https://github.com/stella/folio/pull/914) [`bab04ba`](https://github.com/stella/folio/commit/bab04bafac56c1cb249e1344d2b145574e007c66) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a `w:pict` whose VML shape looks like a watermark but belongs to nobody.

  VML has no serializer, so an unresolved `w:pict` is either replayed from the bytes the parse captured or lost. `shouldPreserveRawVmlPict` declined for any shape `isWatermarkShape` recognised, on the premise that `watermarkParser` had claimed it. That premise was a guess about another module. The watermark reader claims a direct `v:shape` child of a `w:pict`, carrying a non-empty `v:textpath` or a `v:imagedata`, alone in its paragraph, in a header; a `v:oval`, a shape nested in a `v:group`, a shape sharing its paragraph with text and every `w:pict` in a footer or in the body all fall outside it. Those were declined by one owner and claimed by no other, so the artwork was dropped and the relationship its `v:imagedata` named stopped resolving. Verbatim part replay hid it until the part was rebuilt.

  The decline is gone. The watermark owner removes its own artwork from the model, by emptying the paragraph it detached, so the run parser does not have to guess who claimed what; a watermark the header reader does claim is still written once.

- [#927](https://github.com/stella/folio/pull/927) [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Split the preview table by what backs the preview. A source-backed kind keeps the strings the budget matches a `src` by and the cap it charges against; a descriptor-backed kind declares neither, so the diagram entry no longer carries a data-URL prefix nothing starts with, a character cap over no characters, or the mime type and filename the retired raster was stamped with. Only a source-backed kind can be named in a budget override or returned by the matcher.

- [#927](https://github.com/stella/folio/pull/927) [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give a drawing that carries no relationship one spelling, and classify the preview the second spelling hid. `Image.rId` is a `RelationshipId`, a branded non-empty string the parser mints from a real `r:embed`, `r:id` or `r:link`, so the empty string can no longer stand for absence: it reached a save as `<a:blip r:embed=""/>`. A VML shape's render is now preview-only, like the group render beside it, so the editor declines to manipulate it and a save replays the authored `w:pict` instead of writing the render into `word/media/` as the picture the shape had become. Stored collaboration snapshots take both changes through attr-schema version 4.

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve `w:tblPrExChange` like every other tracked property revision, from one list of them.

  The element round-tripped but the editor did not know it: accept, reject and the tracked-change list each carried their own list of four or five change elements, and none of them named the fifth. Accepting every change left the revision on the row, so the document said a formatting change was still pending after the reviewer had resolved it.

  The set is now written down once. `PROPERTY_REVISION_KINDS` is the model's census of the change elements that store a complete previous property set, and one site table says where each one lives, how it resolves and what a reader calls it. The carrier reader, the accept/reject command, the tracked-change list, the comparison's scopes and the Vue sidebar's labels are each total over it, so a revision the model gains is a compile error at every one of those rather than a branch nobody wrote.

  Two revisions the list had already lost come back with it: a paragraph's `w:pPrChange` was read from an attr the schema does not declare, and `w:sectPrChange` was never listed at all.

  Accepting a `w:tblPrExChange` drops the record and keeps the row's current exceptions; rejecting it restores the stored ones wholesale, including restoring their absence.

- [#938](https://github.com/stella/folio/pull/938) [`9d1e897`](https://github.com/stella/folio/commit/9d1e8973631997ee478ca6ad0fedda4cc1134246) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Hold the attribute remainder to one record per block container, a text box included.

  `keepOneAttributeRemainderPerRecord` gives an authored `w:rsid*` set to the first record holding it and to no other, so the half of a split paragraph the editor created claims no revision session. Its walk named paragraphs, tables and content controls and fell through on everything else, and `w:txbxContent` hangs off a shape inside a run rather than off a block child: no record inside a text box was ever reached. Splitting a paragraph or a row in one wrote the source paragraph's `w:rsidR` onto both halves.

  The walk is now `visitBlockTreeRecords`, one traversal that enters a cell, an SDT's content and a shape's text body alike and is exhaustive over the block union, so a new block kind cannot be added without deciding what it holds. `visitDocxParagraphs` is the same traversal with its own pruning.

- [#914](https://github.com/stella/folio/pull/914) [`bab04ba`](https://github.com/stella/folio/commit/bab04bafac56c1cb249e1344d2b145574e007c66) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Re-consolidate a paragraph after a parse-time normaliser removes an inline item.

  `parseParagraph` consolidates a paragraph's runs, and every inline item that is not a run is a merge boundary. Three normalisers then remove items from that already-consolidated array: a comment marker naming a comment the package does not define, a move-range marker with no other half, and the per-paragraph range markers a multi-paragraph comment is cut into. Each left two mergeable runs adjacent, which the parse that consolidated them would never have produced.

  That is an oscillation rather than a loss. Save 1 wrote the pair, the next parse merged it, save 2 wrote one run, so the second save differed from the first. The removal now restores the invariant where it happens, in `InlineContentRemovals.apply`, so every normaliser that removes an inline item gets it and none has to remember.

- [#926](https://github.com/stella/folio/pull/926) [`e14ada1`](https://github.com/stella/folio/commit/e14ada1d3a6aa30a769b9fe21ca7daf847b0a619) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read and write the names ECMA-376 Part 1 spells by writing direction from one generated table.

  Part 1 names a horizontal edge `start`/`end` where Part 4 names it `left`/`right`, and Part 4 declares both so a Transitional consumer reads either. folio rebuilds every package as Transitional, so each of the readers, the writers and the survival law had its own hand-written copy of which two names are one slot.

  `packages/core/src/docx/strictNames.gen.ts` is now that table, derived from the committed schema graph and a cited list of which spelling Part 1 declares. The border, cell-margin and indent readers take both spellings through it rather than falling back name by name, the table and numbering serializers bind the name they write to the key it stands for, and a check fails when a reader hand-lists a spelling again. No output changes: the same equivalences, now generated.

- [#942](https://github.com/stella/folio/pull/942) [`780ecd0`](https://github.com/stella/folio/commit/780ecd0d8d73353bc0d930d97e525d4b516187ac) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a shifted row the two sides hold identically.

  A sole shifted pair that also strands containers on both sides was split back into a deletion and an insertion whenever the container count was unchanged, so a table that lost two rows and gained two reported every row deleted and every row new, the surviving one included.

  That rule is about a mapping inferred from similarity, and a pair whose content digests are equal and whose blocks carry the same ids in the same order is not one: there is nothing left to infer. Both halves are required, since equal content alone is two boilerplate rows reading alike, and equal ids alone are positional ids agreeing after a reorder with no content behind them.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop writing a run the next parse drops, so a package that holds a text box is stable from the first save.

  Whether a run is kept was decided in three places. The parser drops a run that holds no payload, the consolidator dropped a run that held none, and the serializer wrote one anyway. A text box makes the three disagree: the shape is claimed by a second pass over the paragraph, so between the two passes its run is legitimately empty. The consolidator dropped that carrier, which cost the run its `w:rPr`, and when the carrier survived, the save wrote `<w:r><w:rPr…/></w:r>` for it and the next parse dropped that run, so save 2 differed from save 1.

  `runHoldsPayload` now owns the question and the parser's keep rule, the hyperlink walk, the consolidator and the serializer all ask it. The consolidator keeps a payload-less run as the boundary its comment always claimed it was, the serializer writes no run that holds nothing, and a link admits a run on the same terms a paragraph does.

  Two defects the text-box pass hid behind that are fixed with it: it enriches a paragraph before markers carried over from the body are put in front of its content, so the positions it walks are the ones its own `w:p` produced; and filling a carrier now advances its cursor, so a second text box in the same paragraph no longer inserts itself in front of the first.

- [#939](https://github.com/stella/folio/pull/939) [`63c18a8`](https://github.com/stella/folio/commit/63c18a81c3bad202d054f2e7cf32367df9c1c44c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a section break on the paragraph that ends the section.

  A `w:sectPr` inside a `w:pPr` says the section ends at that paragraph's mark. ProseMirror copies a node's attrs when a command splits it, so pressing Enter inside a section-ending paragraph produced two nodes over the same `_sectionProperties` object and the save wrote the break twice: a section nobody added, whose `w:sectPr` repeated the real one's `w:rsidSect` and so claimed its revision history as well.

  The from-leg now assigns the break by reference identity and position — among the paragraphs holding one `_sectionProperties` object, only the last in document order writes it, and the object is never cloned. That is where the rule can be total: Enter, a paste, an AI edit and a split merged in from another client each reach the model by their own route, and only the projection sees them all.

  Joining is the same rule read backwards. A join consumes one paragraph mark and keeps the other, so Backspace at the start of a section-ending paragraph leaves a merged paragraph that still ends the section; ProseMirror keeps the first node's attrs, so the break is now carried to the mark that survived. A mark deleted outright still takes its section with it.

- [#926](https://github.com/stella/folio/pull/926) [`e14ada1`](https://github.com/stella/folio/commit/e14ada1d3a6aa30a769b9fe21ca7daf847b0a619) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a style's indent and cell margins in either spelling of the name.

  ECMA-376 Part 1 spells a horizontal edge `start`/`end` where Part 4 spells it `left`/`right`. The document readers took both, and the style reader — which has its own copies of the same parsers — took only the Transitional one, so a style, a document default or a conditional table region written by a Strict producer lost its indent and its cell margins on the way in. Both now read through the generated rename table, which also covers `w:tblStylePr` and `w:docDefaults`, and a property test over every slot the table names holds every reading site to it.

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a tracked move's `w:name` through the editor.

  `w:moveFromRangeStart` and `w:moveToRangeStart` carry the name that binds a move's source to its destination, and with it the range's own `w:id`, author and date. `toProseDoc` had no node for the four move-range markers and dropped them, so the first save after any edit wrote two unrelated revisions where the document had one relocation.

  The markers are not content a caret can sit in, so they ride on the paragraph beside `bookmarks` and are put back around the wrappers they delimit: a range opens before the first `w:moveFrom` / `w:moveTo` of its kind and closes after the last, and a range that spans several paragraphs still opens in the first and closes in the last. What is carried is the model's own marker, so a field added to `CT_MoveBookmark` is carried without being named again.

- [#913](https://github.com/stella/folio/pull/913) [`94853aa`](https://github.com/stella/folio/commit/94853aa3a42f02fa56079f571ace5368fba3efea) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Report an edited table cell as an edit inside its row, in a document whose
  producer wrote no `w14:paraId`. The row alignment paired a container only when
  its block ids were identical on both sides and their stability had changed from
  positional to stable — the shape a save left behind when it stamped folio's
  minted ids into the package. That made the comparison depend on a side effect
  of the save rather than on the two documents in front of it: with minted ids no
  longer persisted, both sides read positional, the pairing never fired, and one
  edited cell came back as a deleted row plus an inserted row. Container identity
  now goes through `alignParagraphOrdinals`, the same owner the selective save
  asks which paragraph is which, so the two cannot answer differently.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep an empty `w:rPr` on every owner of a run property set.

  `w:rPr` is optional on a run, on the paragraph mark, on a style, on a numbering level and inside either kind of `w:rPrChange`, so a producer that wrote `<w:rPr/>` stated something an absent element does not. The one reader keyed the record on the properties the element yielded and answered `undefined` for one that yielded none, so the empty element reached no model and the one writer put nothing back. The carrier is now the element: the reader returns an empty record for an element that exists, and the writer writes the empty element exactly when the record is present.

  The scan argues for it at every owner. Across the 5335 packages in the public corpus, `<w:rPr/>` appears 5121 times on a run, 3890 on the paragraph mark, 2419 on a style, 674 on a numbering level and 31 inside a `w:rPrChange`, and Word is among the producers of each. On the paragraph mark the case is sharper still: an empty one cannot state formatting, but it is where `w:rPrChange` and the mark's `w:ins`, `w:del`, `w:moveFrom` and `w:moveTo` live, so writing one is pure presence.

  The paragraph mark's emission gains the third answer this needs: `undefined` for a mark that carried no property set, `""` for one that carried an empty one. `ParagraphFormatting.runProperties` is written only by the parser, so a present-and-empty record means the source had the element and nothing else can invent one.

- [#928](https://github.com/stella/folio/pull/928) [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stroke every underline member at the weight it is drawn with. The display list thickened nothing, so `thick` and the seven `*Heavy` members painted on the page and in the PDF as the plain rules they are heavy versions of, while the editor drew them twice as thick: the same run came out two ways. Weight is now a total table over `ST_Underline`, and the heavy multiple is one constant both the stroke and the CSS `text-decoration-thickness` derive from.

  Two members are no longer approximated by what CSS can spell. `wavyDouble` draws two waves rather than two straight rules, and `words` underlines the words and skips the spaces between them: the display list carries one advance per code point, so it can place a span per word where a `text-decoration` cannot.

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read through an inline wrapper in find-and-replace and in the compatibility
  inspector. Both walked paragraph content with an `if`-chain that ended in a
  silent default, so every member neither named was skipped: a phrase inside a
  `w:bdo`, a `w:dir`, a smart tag, a run-level `w:customXml` or a `w:ins` could be
  read on the page and not found, and an opaque drawing inside one of them was
  invisible to the inspector, which then reported a document safe to edit that was
  not. Both walks are now a `switch` with a `never` default, so a content type
  added to the model has to be given a decision. The search projection counts
  every text node the editor's own searchable text counts, deleted text included:
  the offsets it produces are resolved back to an editor position, so a member
  counted on one side and not the other shifts every later offset and the
  replacement lands on the wrong characters.

- [#911](https://github.com/stella/folio/pull/911) [`5a0b642`](https://github.com/stella/folio/commit/5a0b6424bded9be47e878135732506efbb291353) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep omitted OOXML table-grid slots out of rendered and serialized cells.

- [#937](https://github.com/stella/folio/pull/937) [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Nest a leaf's marks in the order the save leg nests its elements.

  Mark rank was the key order of `MARK_EXTENSIONS`; element nesting is decided in `extractParagraphContent`. Nothing held the two together and they had drifted: `hyperlink` was registered before `insertion`/`deletion`, so the editor DOM read `<a><span class="docx-insertion">` while the save wrote `w:ins > w:hyperlink > w:r`, and a rule or a walk written against one nesting was written against a document the other leg does not produce.

  `MARK_NESTING_ORDER` now states the order once, outermost first, and `StarterKit` registers from it; the record keeps its job of naming the marks that exist and building them. Only the marks the save leg gives an element of its own are ranked by it: a comment range around everything a leaf produces, then the revision, then the transparent wrapper the revision takes inside it, then the link and its runs. The rest are run properties with no element in OOXML, so the save leg ranks them against nothing and they stay where presentation put them. A test serializes a leaf under a comment, a revision, a wrapper, a link and bold through both legs and holds their nesting to each other.

  The change is visible in the ProseMirror layer, where an inserted or deleted link is now inside the change's span: the adapter's link colour would paint over the redline, so the anchor inherits it, and the display modes that drop the change colour hand the link colour back.

- [#935](https://github.com/stella/folio/pull/935) [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep one equality over a paragraph's stated numbering.

  The marker tier (`prosemirror/listMarker.ts`) and the layout tier (`layout-bridge/convert/toFlowBlocks.ts`) each carried a private copy of the same two helpers, byte-identical bodies under different names. `isListNumPr` and `sameListNumPr` now live once beside `ListNumPr`, and both copies are gone: an equality that decides whether a list renumbers is not a thing to have two of.

- [#928](https://github.com/stella/folio/pull/928) [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read an underline member's pattern, weight and stroke count off one row.

  The display list held three parallel total tables keyed by `ST_Underline`, so `wavyDouble`'s `wavy` pattern and its second stroke were one decision spelled in two places, and a new member needed three edits. `UNDERLINE_STROKES` states a row per member and `underlinePattern`, `underlineWeight` and `underlineStrokeCount` read it, so no consumer changes.

  The DOM's `text-decoration-thickness` now comes from that row's weight rather than from a second list of the heavy members, which is what kept the page and the editor honest about which members Word draws heavy.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a style's run properties with the same reader as a run's, and resolve a property stated twice to its last statement.

  `styleParser.ts` kept a private copy of `parseRunProperties`, and the copies had drifted. The style copy never read `w:noProof`, wrote an empty `fontFamily` and an empty `color` for a `w:rFonts` or a `w:color` that stated nothing, kept none of the `w:rPr` children folio does not model, and — the difference that changed a value — took the _last_ of two statements of a toggle while the run copy took the first. It also disagreed with itself: `w:rtl`, `w:cs` and `w:dstrike` took the first statement while `w:b` and `w:strike` took the last.

  `EG_RPrBase` is an `xsd:choice` referenced `maxOccurs="unbounded"`, so a repeat is valid markup rather than a malformed file: 44 of the 5299 packages in the public corpus hold one. The last statement wins, and the statements it beat are not written back, so the saved element states each property once and a consumer that takes the first and one that takes the last read the same value from it. The evidence is recorded as `repeated-run-property-resolves-last`.

  A style's `w:rPr` now also keeps what folio models nothing for, the way a run's already did.

- [#934](https://github.com/stella/folio/pull/934) [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write a compiled `w:rPr`'s children in the schema's order, from the same generated list folio-core writes from.

  `@stll/docx-core` holds a second `w:rPr` writer — the one the legal-source compiler and the build-from-scratch export share — and it had grown an order of its own, emitting `w:highlight`, `w:sz` and `w:szCs` ahead of `w:rFonts`. A run carrying both a font and a size therefore came out in one order from this package and another from folio-core's serializer. `EG_RPrBase` is an `xsd:choice` referenced `maxOccurs="unbounded"`, so both spellings are valid; what the canonical order buys is one form, the one Word writes, from both writers.

  The generated order moves down to where both can read it: `@stll/docx-core/schema` is a new subpath exporting `SEQUENCE_CHILDREN` and the writer that orders by it, and folio-core's declared-child table spreads that same object in. One emitted order, one sort, and a serializer that cannot restate either.

- [#933](https://github.com/stella/folio/pull/933) [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write one spelling of an `ST_OnOff` value: `0` for an off, the bare element for an on.

  `w:hideMark` wrote its explicit off as `w:val="off"` while every other on/off element in the package wrote `0`, and `w:updateFields` wrote its on as `w:val="true"` while its neighbour in the same part wrote the bare element. All three spellings mean the same thing, so a package that mixes them only makes every byte-level comparison argue about which one it is looking at.

  `scripts/on-off-spelling.test.ts` holds the tree to the one spelling, with no allowlist. Reading is unchanged: `parseOnOffValue` and `parseBooleanElement` take all six spellings, and a captured element still replays the bytes it arrived as.

- [#928](https://github.com/stella/folio/pull/928) [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Type a shape or text-box outline's dash as `ST_PresetLineDashVal`, and give each stroke vocabulary its own table.

  `ShapeOutline.style` was a hand-written eleven-member union named after CSS, holding what `a:ln/a:prstDash@val` declares. It is now `ShapeOutline.dash`, typed with `PresetLineDashVal`, generated from the committed schema graph by `scripts/generate-preset-line-dash.ts` with the same write/check pair `BorderStyle` uses; `bun run generate:preset-line-dash:check` runs in CI. A `@val` the schema does not declare is kept as `{ kind: "unrecognised", raw }`, written back unchanged, and reported through `ParseContext` as `outline-dash-outside-enum`. `a:custDash` is a different element and stays unmodelled: an outline that carries one replays through `ShapeOutline.rawXml`.

  The display list resolved three vocabularies through one lookup keyed by lower-cased strings: a CSS `border-style`, a DrawingML preset dash, and a CSS `text-decoration-style`. `dash`, `dot` and `solid` collide across them, and every member no other vocabulary spells the same way had no entry and painted as a plain line. Nine of the eleven preset dashes (`dot`, `lgDash`, `dashDot`, `lgDashDot`, `lgDashDotDot`, `sysDash`, `sysDot`, `sysDashDot`, `sysDashDotDot`) and the seven heavy underline members were in that set, so a `sysDash` outline and a `dottedHeavy` underline both stroked solid. There are now three tables, each `as const satisfies Record<Union, StrokePattern>` over its own vocabulary, and each consumer calls the one it speaks.

  The DOM painter had the same defect one step further on: it interpolated the outline's dash straight into a CSS `border` shorthand, so `border: 2px sysDash #000` was invalid and a dashed text-box outline did not paint at all. A dash is now translated to a CSS keyword before it reaches a shorthand, and `run.underline.style` is translated rather than assigned, which is what made `text-decoration-style: dottedHeavy` a no-op.

- [#942](https://github.com/stella/folio/pull/942) [`780ecd0`](https://github.com/stella/folio/commit/780ecd0d8d73353bc0d930d97e525d4b516187ac) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Publish only the files consumers load: `dist`, the `src` the `exports` map resolves to, and the licence and notice texts. Tests, snapshots, fixtures, build scripts and `tsconfig` files no longer ship.

- [#928](https://github.com/stella/folio/pull/928) [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render an underline in the editor from the same table the painters read.

  `UnderlineExtension.toDOM` carried a private four-entry table (`double`,
  `dotted`, `dash`, `wave`) while the display list and the DOM painter read the
  total one, so the eleven members it had never heard of drew a plain line in the
  editor and their own pattern on the page: `dottedHeavy`, `dashLongHeavy`,
  `wavyHeavy` and `words` among them. The table is gone. `toDOM`, the DOM painter
  and `textToStyle` now all render a member through `underlineDecorationCss`,
  which is `as const satisfies Record<UnderlineStyle, …>`, so a member added to
  the enumeration cannot reach a backend without a decision attached.

  The one table also states the weight CSS has no keyword for: `thick` and the
  seven `*Heavy` members carry a `text-decoration-thickness` of twice the ratio
  the display list strokes a plain underline with, so both DOM backends scale it
  with the font size. The remaining approximations are recorded beside the table:
  `words` underlines the spaces between words (`text-decoration-skip-ink` skips
  descender ink, not spaces, and `text-decoration-skip: spaces` never shipped),
  `wavyDouble` draws two straight lines, and `dashLong`, `dotDash` and
  `dotDotDash` draw the single dash pattern CSS has.

  `toDOM` states the line and the style in one `text-decoration` shorthand, which
  the mark's existing parse rule reads back through the table's inverse: each CSS
  keyword resolves to its canonical author, the plain member drawn exactly that
  way (`solid` to `single`, `dashed` to `dash`, `wavy` to `wave`). The table is
  many-to-one, so what CSS cannot spell does not survive the round trip:
  `dottedHeavy` parses back as `dotted`, `words` and `thick` as `single`. A value
  naming no keyword folio writes parses as the plain underline, and a
  `text-decoration-style` is not read on its own, which would make a dotted
  strikethrough an underline.

  `w:u w:val="none"` now cancels an inherited underline in the DOM painter as it
  already did in the display list: the member carries no declarations, so the
  painter draws no line rather than an unstyled one.

- [#939](https://github.com/stella/folio/pull/939) [`63c18a8`](https://github.com/stella/folio/commit/63c18a81c3bad202d054f2e7cf32367df9c1c44c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a save that adds a section the editor never authored.

  The repack fidelity guard read one direction of the section count: it refused a save that dropped a section and said nothing about one that added a section, which is why a split paragraph's duplicated `w:sectPr` reached the file without anything noticing.

  A gain is legitimate, since the editor inserts breaks, so the guard asks what the model holds rather than what the original held. Two paragraphs over the same `SectionProperties` record are one section's split halves, never two sections: that fails with `DocxDuplicateSectionCarrierError`, which names the paragraph. The package must also state exactly as many `w:sectPr` elements as the model has records, because a record the serializer fails closed on would otherwise pass as a smaller gain rather than a loss.

- [#911](https://github.com/stella/folio/pull/911) [`5a0b642`](https://github.com/stella/folio/commit/5a0b6424bded9be47e878135732506efbb291353) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match Word's terminal font fallback for imported stories, keep tracked deletions on their original metrics, and render deterministic review bars.
- Updated dependencies [[`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19), [`20c9522`](https://github.com/stella/folio/commit/20c95224287ee8b7bfcb57defc17cac56a54fbcf), [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480), [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d), [`9d1e897`](https://github.com/stella/folio/commit/9d1e8973631997ee478ca6ad0fedda4cc1134246), [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a), [`0249b81`](https://github.com/stella/folio/commit/0249b811d88f80aa2d3d2bf2c0d1a9c20c8dc82d), [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d), [`0249b81`](https://github.com/stella/folio/commit/0249b811d88f80aa2d3d2bf2c0d1a9c20c8dc82d), [`0249b81`](https://github.com/stella/folio/commit/0249b811d88f80aa2d3d2bf2c0d1a9c20c8dc82d), [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d), [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d), [`3fde3f4`](https://github.com/stella/folio/commit/3fde3f47fcdf038dd4a787a65fb46c4342752da1), [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a), [`b6a1a58`](https://github.com/stella/folio/commit/b6a1a58d508c7296df2b6bd52ed1ae48b58cde1a), [`3fde3f4`](https://github.com/stella/folio/commit/3fde3f47fcdf038dd4a787a65fb46c4342752da1), [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f), [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f), [`3b984e5`](https://github.com/stella/folio/commit/3b984e5759f40fe5af1c658baa9c163078b7db69), [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f), [`9d1e897`](https://github.com/stella/folio/commit/9d1e8973631997ee478ca6ad0fedda4cc1134246), [`48959b9`](https://github.com/stella/folio/commit/48959b925274492499f2ba85097d77e09c50d53f), [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3), [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19), [`bba0c3c`](https://github.com/stella/folio/commit/bba0c3c7b88b248a27b2096baed46dcc60a78d2d), [`166d3f0`](https://github.com/stella/folio/commit/166d3f0868393dc74c0609ac4a94cba0e579743b), [`166d3f0`](https://github.com/stella/folio/commit/166d3f0868393dc74c0609ac4a94cba0e579743b), [`56539e3`](https://github.com/stella/folio/commit/56539e3504cf2ed26e6b2e016bd66507581d5c24), [`d098792`](https://github.com/stella/folio/commit/d098792ef0634154446a13d80757b7f73b233838), [`fa2abc1`](https://github.com/stella/folio/commit/fa2abc134692faf7dd48cdeb28d0631ee9a796b7), [`b4dce7a`](https://github.com/stella/folio/commit/b4dce7ae31fa92b7a2ae8f3f5c2286bce2822e19), [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3), [`8f9a01b`](https://github.com/stella/folio/commit/8f9a01b2b849e743a3dbac019e7a1aecbf9c2379), [`f1a4d2d`](https://github.com/stella/folio/commit/f1a4d2dd55fc83bc3872253fd4b00a017785ec85), [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3), [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3), [`3b984e5`](https://github.com/stella/folio/commit/3b984e5759f40fe5af1c658baa9c163078b7db69), [`8f26a09`](https://github.com/stella/folio/commit/8f26a09ca39f764ac2d99bb3ad8dd01a3377e7d3), [`edbc88b`](https://github.com/stella/folio/commit/edbc88b4dbc4a0958b84f9d36859042f0c9489e3), [`13d3f50`](https://github.com/stella/folio/commit/13d3f50278f2d6cb9553004b38a0ebe0f611f480), [`da1fc6a`](https://github.com/stella/folio/commit/da1fc6a4c8d7a48705187d164d8e8180cefeff74)]:
  - @stll/docx-core@0.25.0

## 0.46.0

### Minor Changes

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every block-level child folio does not model, where it stood, through the editor as well as through a save.

  `w:body`, `w:hdr`, `w:ftr`, `w:tc`, an SDT's content and a note body share one walk, and it modelled paragraphs, tables and content controls and let the rest fall off the end. A `w:permStart` between two paragraphs is the whole of a document-protection range; `w:altChunk` is an entire imported document; `m:oMathPara` is a display equation. The walk now goes through the shared child dispatcher, whose handler map the compiler makes total over the children the schema declares for a block container and whose default is the verbatim sink.

  `BlockContent` gains a `preservedBlock` member holding the captured markup, and the editor gains a zero-width `preservedBlock` node for it. Position is structural on both sides: the capture sits between the same two blocks in the model, in the ProseMirror document and in the saved part, so inserting, splitting or deleting a neighbour moves it the way a reader would expect and nothing has to keep an index honest.

  `Paragraph`, `Table` and `BlockSdt` lose `rawMarkersBefore` / `rawMarkersAfter`, the narrower mechanism this replaces: it kept only sixteen range-marker names, dropped them when the container held no block at all, and had no editor leg, so a document that survived an untouched save lost the markup the moment anybody opened it. `Footnote.content`, `Endnote.content` and `TableCell.content` are now `BlockContent[]` rather than hand-written copies of it.

- [#885](https://github.com/stella/folio/pull/885) [`280a9be`](https://github.com/stella/folio/commit/280a9be8d908cd74108c96a085679347918ea712) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry `w:commentReference` through the editor as an inline node of its own, so
  a comment's visible mark keeps the place it was authored in. The mark-only
  projection recorded no position for it and the serializer guessed one after
  every range end: two comments closing together came back interleaved rather
  than grouped, and a comment spanning three paragraphs came back marked three
  times. The serializer now writes what the model says, the schema gains a
  zero-width `commentReference` node whose integrity is repaired on every
  transaction, and a model that arrives with a range and no reference is
  completed once for the whole story rather than per paragraph.

- [#899](https://github.com/stella/folio/pull/899) [`036fcf2`](https://github.com/stella/folio/commit/036fcf2a38c9fadaefdfb46cf9a37a326e196ec7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A drawing's authored rotation and flips survive the editor projection. The
  editor carried `a:xfrm`'s three values inside one CSS transform string, which
  can state neither `rot="0"` nor `flipH="0"`: both spell the identity, which is
  what an absent transform already means, so opening a document and saving it
  again dropped the attribute. The `image`, `shape` and `textBox` nodes now carry
  each value explicitly (`docxRotation`, `docxFlipH`, `docxFlipV` on `ImageAttrs`,
  `ShapeAttrs` and `TextBoxAttrs`; `null` for absent) and write it back, while the
  CSS string stays a projection of them for rendering. A rotate or flip from the
  editor states every value it decides, so rotating back to zero says zero rather
  than handing the decision back to the file.

  The attrs are additive: a node persisted without them is read from its CSS
  string as before, which is the only record such a node has. The collaboration
  attr schema still goes to version 3, because a snapshot a newer build wrote
  must not reach an older one that would drop the three keys unread.

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

- [#912](https://github.com/stella/folio/pull/912) [`94a12dd`](https://github.com/stella/folio/commit/94a12ddd1a959d98e94f5ddd4c1d81276261f805) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write `w:bdo`/`w:dir` back around the text the editor still says they hold.

  The projection lifts a transparent inline wrapper out of the paragraph's content tree and records the nesting on the `inlineWrapper` mark of the leaves it held. The save leg ignored that mark, so a wrapper only survived where the source paragraph's markup was replayed and an edited span lost it.

  `fromProseDoc` now cuts the paragraph's inline sequence into maximal groups of equal stack before it builds runs, and closes the wrappers around each group. A revision stays outermost — `w:ins > w:bdo > w:hyperlink > w:r` — because folio already writes a revision outside the hyperlink it spans, the parse leg is revision-owned, and accepting or rejecting one is a range operation over the revision's own content. A group with nothing left in it writes no wrapper, so a wrapper whose text was deleted or rejected disappears with it.

  A paragraph that was not edited keeps its authored markup, including a wrapper the author put outside a revision: selective save replays its bytes. Rebuilt from the editor, that order is written the canonical way round.

- [#901](https://github.com/stella/folio/pull/901) [`051dbd6`](https://github.com/stella/folio/commit/051dbd612dc6541df1725a29d7bfea8612bed056) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the attributes an element carried that folio has no field for, through a save and through the editor.

  Word writes a revision-session id on nearly every paragraph, run, row and section (`w:rsidR`, `w:rsidRPr`, `w:rsidDel`, `w:rsidP`, `w:rsidRDefault`, `w:rsidTr`, `w:rsidSect`). folio rebuilt each of those elements from the model alone, so opening a document and saving it rewrote the whole revision history.

  `Paragraph`, `Run`, `TableRow` and `SectionProperties` gain `preservedAttributes`, an ordered list of resolved `{ namespace?, name, value }` records. The decision of what to keep is made on the resolved namespace URI and local name, so a source that binds a second prefix to the WordprocessingML namespace does not get a second copy of an attribute the parser already read; the writer is handed the modelled attributes it is about to emit and drops any remainder entry that would spell one of them again, so a duplicate attribute cannot reach the part. A namespace declaration is never in the remainder, and neither is an attribute whose namespace the rebuilt part cannot bind.

  The remainder follows the record: a paragraph, row or section the editor creates from scratch has none, an authored one's survives `toProseDoc`/`fromProseDoc` unchanged, and when a command splits a record in two the half that comes first in document order keeps it. A run has no record in the editor — it is text plus marks — so a run's remainder survives a save and not the projection.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every inline child folio does not model, at its source position. One walk serves a paragraph, the four run-level tracked-change wrappers, `w:bdo`/`w:dir` and an inline content control, and it now goes through the shared child dispatcher over a handler map the compiler makes total. `ParagraphContent` gains a `preservedInline` member holding the captured markup, so `w:permStart`, `w:proofErr`, `w:customXml`, the eight custom-XML revision ranges and `w:subDoc` survive a save and the editor round trip.

  Inside a tracked change the position is the point: markup lifted out of a `w:ins` is markup the reviewer no longer accepts or rejects with the change, so the capture sits inside the wrapper in the model, in the editor and in the saved part. `w:customXml` also keeps the text it puts on the line.

  A bare OMML element is now read by namespace rather than by falling off the end of a switch, and `m:oMathPara` keeps its display form.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every run child folio does not model instead of letting it fall off the end of the run-content switch. `RunContent` gains a `preservedXml` member holding the captured markup at its source position, plus the visible text it contributes, so `w:ruby`, `w:contentPart`, `w:pgNum`, `w:annotationRef`, the note markers and any foreign or future element survive a save and read as text.

  The keep rule now asks the model rather than the source element. The two disagreeing was a two-save oscillation rather than a loss: the first save wrote a run whose payload the model never held, the next parse dropped that run, and the second save differed from the first.

- [#889](https://github.com/stella/folio/pull/889) [`5f0c65b`](https://github.com/stella/folio/commit/5f0c65bd00d49d79ca2cafb0a31efb411e9bab4c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Version the node attrs Folio persists in a collaboration document. A snapshot written by a newer attr schema is now refused with a typed error instead of being rebuilt attr by attr, and `migrateFolioYjsSnapshot` carries a stored snapshot forward offline.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove the verbatim sink's attribute remainder. `PreservedMarkup.attributes`, the `PreservedAttribute` type, the dispatcher's `modelsAttribute` option and `serializePreservedAttributes` had no caller in the product: no container ever passed the predicate, so no attribute was ever kept, and the shape read as coverage that was not there. `docs/container-contract.md` records the design and what wiring it needs.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add the ordered verbatim sink and the shared child dispatcher, and put `w:comment` bodies on them.

  `PreservedMarkup` holds a container's unmodelled children with their position relative to its modelled ones, plus an ordered attribute remainder, so the serializer puts them back between the same siblings rather than at the end. `dispatchChildren` walks a container with a handler map the compiler makes total over the children the schema declares for it, and routes anything undeclared — a foreign namespace, an `mc:` construct, an element a later OOXML revision adds — to the sink by default.

  A comment body may hold everything a document body can. folio modelled only `w:p`, so a table, an equation, a content control, a bookmark or a range marker in a reviewer's comment disappeared on save; `Comment.preserved` now keeps them.

### Patch Changes

- [#892](https://github.com/stella/folio/pull/892) [`9035639`](https://github.com/stella/folio/commit/90356394a8d68e078bbaa95ad2e6643fcff51a1f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `w:cols/@w:sep` as the section authored it. The parser recorded only `true` and the serializer emitted only `true`, so an explicit `w:sep="0"` was read as an absence and written back as one. The attribute was also missing from `serializeColumns`' bail-out condition, so a `w:cols` whose only stated setting was the separator lost the whole element rather than the one attribute. `@w:equalWidth`, which already round-tripped correctly, joins it in the reserved-value registry so both column toggles are recorded against the reader that owns them.

- [#885](https://github.com/stella/folio/pull/885) [`280a9be`](https://github.com/stella/folio/commit/280a9be8d908cd74108c96a085679347918ea712) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a comment range whole across the paragraphs it covers. The conversion
  tracked the open ranges per paragraph, so a comment on paragraphs 1 to 3 marked
  the first and the last (where its two boundaries sit) and left the middle
  unhighlighted, and the save path, reading each paragraph on its own, wrote one
  range per marked paragraph where the author wrote one. The open ranges now flow
  with the block walk, into table cells, text boxes and content controls, and a
  save emits one `w:commentRangeStart` at a comment's first marked position and
  one `w:commentRangeEnd` at its last.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match a container's declared children by namespace as well as local name in the shared child dispatcher. A child from another namespace now reaches the sink instead of the handler its local name happens to collide with, so `m:r` is no longer read as a text run, emptied and pruned.

- [#892](https://github.com/stella/folio/pull/892) [`9035639`](https://github.com/stella/folio/commit/90356394a8d68e078bbaa95ad2e6643fcff51a1f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `@w:fldLock` and `@w:dirty` as a field authored them, on `w:fldSimple` and on the `w:fldChar` that opens a complex field. Three readers and three writers had each collapsed the two attributes to "present and true", so an explicit `w:dirty="0"` -- a field inside a `TOC` result that says not to recompute -- parsed as an absence and saved as one. Both directions now live in one module, `docx/fieldState`. The editor keeps the distinction too: the field node's `fldLock` and `dirty` attrs default to absent rather than `false`, so projecting a field through the editor no longer invents an explicit off. The attributes are written as `1`/`0`, matching Word. Because the persisted attr shape changes, the collaboration attr schema goes to version 2, and `migrateFolioYjsSnapshot` drops the `false` a version-1 snapshot stored for a field that authored neither flag.

- [#890](https://github.com/stella/folio/pull/890) [`7245027`](https://github.com/stella/folio/commit/7245027781760b14d24df47af5ae9f9531be795a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop inventing a drawing's `a:graphicFrameLocks`. Rebuilding a picture whose model held no lock record wrote `noChangeAspect="1"`, because one `undefined` meant both "the author wrote no frame properties" and "folio created this picture": at serialization the two are indistinguishable, so an edited document gained a lock its source never carried. The default now belongs to the insert that creates a picture, and the serializer writes the element only when the model holds locks.

- [#904](https://github.com/stella/folio/pull/904) [`0ebf68f`](https://github.com/stella/folio/commit/0ebf68fb746132181407783d524fb39423ee8b5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Refuse a note-part patch that would leave a comment range with only one half.
  A comment can be anchored on a footnote's or endnote's own text, so its range
  spans that note's paragraphs; splicing only the paragraph an edit touched then
  wrote the other half alone, which is invalid OOXML and anchors the comment to
  nothing. The refusal now belongs to the one splice primitive every selective
  patch goes through, and a refused note part is rewritten whole from the model
  instead of failing the save.

- [#890](https://github.com/stella/folio/pull/890) [`7245027`](https://github.com/stella/folio/commit/7245027781760b14d24df47af5ae9f9531be795a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a `paraId` by namespace URI wherever it is written. A paragraph's id, a `commentsExtensible` join key and a `commentsExtended` thread link were each resolved by prefix, with a local-name fallback that matched any prefix at all: a file binding `w14`, `w15` or `w16cex` to a prefix of its own was read correctly only by luck, and an unrelated `vendor:paraId` was read as a thread key, carrying another thread's date, parent and resolved state into the comment. One reader now answers "the paraId of this element" for all of them, and it accepts the Word 2010, 2012 and 2018 namespaces and no others.

- [#906](https://github.com/stella/folio/pull/906) [`34ee586`](https://github.com/stella/folio/commit/34ee58684fd0700088e92deae7900198878c6849) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read through a bidirectional wrapper in the save-side resource census and the
  rendered-page-break detector. `w:bdo` and `w:dir` are transparent, and both
  passes stopped at one: a hyperlink authored inside a wrapper got no `r:id`,
  which is the whole of how an `href` is saved, so the package held a link
  pointing nowhere; and a `w:lastRenderedPageBreak` under one was invisible to
  the detector that decides where the break is re-emitted.

- [#891](https://github.com/stella/folio/pull/891) [`31ab1fa`](https://github.com/stella/folio/commit/31ab1fa4884cbd837933028e5c1f1ecf8324254e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a paragraph's `w14:textId` by namespace URI. A prefix is an alias, and the prefix lookup fell through to an any-prefix local-name match, so a `textId` bound to a foreign namespace was taken for Word's paragraph identity and written back as one. An alternate prefix bound to the Word 2010 URI still reads.

- [#891](https://github.com/stella/folio/pull/891) [`31ab1fa`](https://github.com/stella/folio/commit/31ab1fa4884cbd837933028e5c1f1ecf8324254e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a run inside overlapping comment ranges in every comment it belongs to. The painted run advertised only the first id, so hover and active styling, and the sidebar anchor, answered for one comment and denied the other. A run in more than one range now also carries `data-comment-ids`, the whole membership, and the adapters read the painted anchors through one shared helper rather than the first id alone; a run in a single range paints exactly what it painted before.

- [#890](https://github.com/stella/folio/pull/890) [`7245027`](https://github.com/stella/folio/commit/7245027781760b14d24df47af5ae9f9531be795a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop `word/styles.xml` growing by a copy on every save. Which styles the original part already defines was decided by scanning its text for `w:styleId="…"`, and that text is the XML spelling of an id while the model holds its decoded value: a single-quoted attribute, an id carrying an escaped character such as `Header &amp; Footer`, or an earlier attribute whose value contains `>` all read as a style the part lacked, so it was appended again each time the document was saved. The part is now read as XML, and a style id is written at most once.

- [#891](https://github.com/stella/folio/pull/891) [`31ab1fa`](https://github.com/stella/folio/commit/31ab1fa4884cbd837933028e5c1f1ecf8324254e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep an authored image or shape transform through a save, including zero. `rot="0"` and `flipH="0"` are OOXML's defaults, so the truthiness guard on `a:xfrm` could not tell an authored zero from an absent attribute and the save dropped it; the image parser also read `rot="0"` as no rotation at all. Rotation and both flips are now read as authored values and written iff the model holds one.

- [#885](https://github.com/stella/folio/pull/885) [`280a9be`](https://github.com/stella/folio/commit/280a9be8d908cd74108c96a085679347918ea712) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the annotations a replaced span carried. ProseMirror drops a mark declared
  `inclusive: false` once a replacement reaches the end of what the mark covers,
  which is right for formatting and wrong for a mark that names something outside
  itself. Replacing a block's text therefore dropped its `comment` marks — taking
  the comment's only range start with them, so the save wrote a `commentRangeEnd`
  with no start, invalid OOXML and a comment anchored to nothing — and dropped
  its `hyperlink` mark, leaving prose that had been a link pointing nowhere.

  A replacement now carries the comments and the link its span held (in tracked
  mode too, so accepting the change keeps them), keeps the zero-width anchors
  inside it (comment references, bookmark boundaries, text-box anchors) rather
  than deleting them, and a selective save refuses to patch a story whose comment
  ranges it would leave with one half, falling back to a full repack instead.
  Every non-inclusive mark now has a recorded disposition, so a new one cannot
  join the schema without a decision.

- [#890](https://github.com/stella/folio/pull/890) [`7245027`](https://github.com/stella/folio/commit/7245027781760b14d24df47af5ae9f9531be795a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep an empty header or footer empty. The rebuild path added `<w:p><w:pPr/></w:p>` whenever the part came out with no blocks, on the premise that OOXML requires one: `CT_HdrFtr` holds a single `EG_BlockLevelElts` occurrence whose choice members are all optional, so a part with no block children is valid and is what Word writes for a blank header. Verbatim replay returned such a part unchanged, so the invented line appeared only after the document had been edited.

- [#892](https://github.com/stella/folio/pull/892) [`9035639`](https://github.com/stella/folio/commit/90356394a8d68e078bbaa95ad2e6643fcff51a1f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write `CT_Border/@w:shadow` and `@w:frame` when the border authored them, whichever state they authored. The serializer emitted the attribute only when the model held `true`, so an explicit `w:shadow="0"` came back from a save as an absence on every border position (paragraph, style, table, cell and page). Neither attribute carries an XSD default, so the two are not interchangeable. `parseOnOffAttribute` already kept all three states; only the emit collapsed them. The value written is now `1`/`0`, matching Word and the table and section serializers, rather than `true`.

- [#890](https://github.com/stella/folio/pull/890) [`7245027`](https://github.com/stella/folio/commit/7245027781760b14d24df47af5ae9f9531be795a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Export a document containing an interlaced PNG to PDF. The PDF image decoder refused Adam7 outright, alongside the 16-bit samples it cannot represent exactly, so a document Word displays without comment failed to export at all. Adam7 is lossless and exactly representable — seven ordinary filtered rasters of the same samples — so it is now decoded: each pass is unfiltered and scattered into the full raster, and the decompression budget counts the passes rather than assuming progressive geometry. A 16-bit PNG is still refused, interlaced or not.

- [#900](https://github.com/stella/folio/pull/900) [`30ced66`](https://github.com/stella/folio/commit/30ced66282baa9f7d494206c7aca7aeabf08c98f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Run the collaboration attr-schema migration on every path that reads a stored fragment, not only in the offline sweep. A snapshot written under an older attr schema reached `initProseMirrorDoc` unmigrated, so a value a step rewrites was read in its old shape as the new one; and the marker was stamped only when Folio happened to rewrite the whole fragment, so an editor could write this build's attrs into a fragment still marked older, which an older build would then read and drop unnoticed. `applyAttrSchemaMigrations` owns both, and the editor, the server materialization and `migrateFolioYjsSnapshot` all go through it.

  A rotate from the editor states the turn the drawing is at. `%` keeps the sign of its left operand in JavaScript, so rotating a drawing whose authored `rot` was negative answered with a negative rotation.

- [#898](https://github.com/stella/folio/pull/898) [`74a9ac4`](https://github.com/stella/folio/commit/74a9ac45acb03be7e7d19939e99806a101eceb91) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve pending native selections when focus is requested redundantly.

- [#902](https://github.com/stella/folio/pull/902) [`ff50230`](https://github.com/stella/folio/commit/ff50230e98a50efaeca00674f2ef9b7be90ddc03) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored Word symbols when checkbox content controls change state.

- [#903](https://github.com/stella/folio/pull/903) [`30b7e4c`](https://github.com/stella/folio/commit/30b7e4c234b0e68abaec1e1bd1cff11bbcca672a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve direct text formatting after returning to a paragraph created with Enter.

- [#906](https://github.com/stella/folio/pull/906) [`34ee586`](https://github.com/stella/folio/commit/34ee58684fd0700088e92deae7900198878c6849) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read through a bidirectional wrapper when exporting markdown. `w:bdo` and
  `w:dir` say how their text is laid out, not what it is, and both inline
  renderers narrowed by a switch whose default contributed nothing: a paragraph
  whose runs sat inside one exported as an empty line, in the pipe-table path and
  the HTML-cell path alike. The wrapper is the ordinary way to write a
  right-to-left run, so the loss fell entirely on right-to-left documents.

- [#889](https://github.com/stella/folio/pull/889) [`5f0c65b`](https://github.com/stella/folio/commit/5f0c65bd00d49d79ca2cafb0a31efb411e9bab4c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep body text out of the PDF outline. A paragraph whose `w:outlineLvl` is the reserved value 9, such as a `TOC Heading`, is no longer written as a bookmark nested nine levels deep.

- [#888](https://github.com/stella/folio/pull/888) [`4f8eed6`](https://github.com/stella/folio/commit/4f8eed643be10173a177695b343e36e2d02d98cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Represent a `w:pict` once. A legacy VML group carrying a text box was claimed by two owners: the run parser kept the whole `w:pict` as one raw drawing, and the text-box pass rebuilt its first `v:textbox` as an editable shape beside it. A save wrote both, so the box's text appeared twice in the saved document, and twice again on every later save. The text-box pass now asks the run parser's own predicate whether a pict is already claimed instead of re-deriving the answer from the markup.

- [#897](https://github.com/stella/folio/pull/897) [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record `Comment.preserved` in the reserved-value registry, whose field-level totality gate caught it, and note on `serializeDocumentToDocx` that it writes a built document rather than a parsed one.

- [#888](https://github.com/stella/folio/pull/888) [`4f8eed6`](https://github.com/stella/folio/commit/4f8eed643be10173a177695b343e36e2d02d98cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open every document that carries an explicit page-break run. `w:br w:type="page"` is an ordinary run child, so Word writes one in a bordered, framed or outlined paragraph, inside a table cell or a text box, and beside any inline kind. Folio refused several of those shapes at conversion and again at layout, which meant the document could not be opened in the editor, laid out or exported to PDF at all. They now project, save and round-trip; where layout can only approximate the break's owner, it says so through the parse-warning channel under the new `page-break-projection-approximated` code instead of throwing. `UnsupportedDocxToProseMirrorConversionError` goes with the last refusal that raised it.

- [#896](https://github.com/stella/folio/pull/896) [`fab7746`](https://github.com/stella/folio/commit/fab774635d347da2337ec99522b21316b65911d8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Dismiss content-control pickers when the user presses outside them.

- [#893](https://github.com/stella/folio/pull/893) [`ab8e3c4`](https://github.com/stella/folio/commit/ab8e3c45d8ddaaa5dfc175a98aed047c86bd003c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Load DOCX packages whose XML parts use UTF-16 little-endian or big-endian encoding.

- [#907](https://github.com/stella/folio/pull/907) [`7be6a09`](https://github.com/stella/folio/commit/7be6a09d11a69c78e3d42bb3418ba5cb05717755) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and render DrawingML image brightness and contrast.

- [#888](https://github.com/stella/folio/pull/888) [`4f8eed6`](https://github.com/stella/folio/commit/4f8eed643be10173a177695b343e36e2d02d98cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a result-less `PAGE` or `NUMPAGES` field result-less on save. The save path wrote a literal `1` into the result of a field the author left empty, so a document reopened from folio said "1" where the source said nothing, whatever page the field sits on. Layout computes the number from the page it paints, so the invented result added nothing and changed what the document says. `proseDocToBlocks` no longer takes an `emptyFieldResult` mode and `EmptyFieldResultMode` is no longer exported: there is one behaviour now.

- [#905](https://github.com/stella/folio/pull/905) [`1f8c6d8`](https://github.com/stella/folio/commit/1f8c6d86f5f0fdcab4a21119ea1436171929c3fa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve rich clipboard content when a rendered image is present beside it.

- [#886](https://github.com/stella/folio/pull/886) [`5bdebe6`](https://github.com/stella/folio/commit/5bdebe62e18908d4c84f6a39b0fa17baf09f417e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Say why a comparison could not be saved. `CompareDocxSerializeError` carried the reason in `cause` and a constant sentence in `message`, so every distinct save failure read identically in a log, a report or a census and none of them could be told apart without a debugger. The message now names the underlying failure, and `cause` still carries it structured.

- [#904](https://github.com/stella/folio/pull/904) [`0ebf68f`](https://github.com/stella/folio/commit/0ebf68fb746132181407783d524fb39423ee8b5c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route the remaining part patchers through the splice owner: stamping paragraph
  ids and restoring a numbering level's custom format both cut regions out of a
  serialized part by hand, so a comment range crossing one of those regions could
  lose a half. A lint rule now holds the boundary.

- [#910](https://github.com/stella/folio/pull/910) [`2eca741`](https://github.com/stella/folio/commit/2eca74133e1c372c89ede5250f6befaf2d27297e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read maths in an ISO Strict package as an equation, and keep an anchored drawing's host paragraph attributes.

  The child dispatcher looks a namespace-keyed disposition up by the URI the element carries, and the dispositions are written in Transitional. Strict spells the maths namespace `purl.oclc.org/ooxml/officeDocument/math`, so a Strict document's `m:oMath` missed the lookup and went to the verbatim sink: the bytes survived and the equation stopped being one, with nothing left to render, edit or read text from. Both the dispatcher and the reader behind the disposition now resolve the namespace through the generated Strict/Transitional pair table, so the rule holds for every namespace-keyed disposition rather than for maths alone.

  Word writes a floating shape into a paragraph of its own. folio lifts the shape out as a block node and drops that paragraph from the editor projection, so the attributes the `w:p` carried and the model has no field for, `w:rsidR` and its family, had no carrier and were gone on the way back. The text box node stands in for the host paragraph and now carries the host's remainder, which the save leg puts back on the paragraph it rebuilds; only the first node of a group takes it, because one paragraph is rebuilt for the group.

- [#895](https://github.com/stella/folio/pull/895) [`9a66b31`](https://github.com/stella/folio/commit/9a66b31dabd64291cf7519c0a16acc4035cf2eed) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve direct run formatting when splitting a paragraph at its end.
- Updated dependencies [[`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697), [`9035639`](https://github.com/stella/folio/commit/90356394a8d68e078bbaa95ad2e6643fcff51a1f), [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697), [`34ee586`](https://github.com/stella/folio/commit/34ee58684fd0700088e92deae7900198878c6849), [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697), [`051dbd6`](https://github.com/stella/folio/commit/051dbd612dc6541df1725a29d7bfea8612bed056), [`817c3cf`](https://github.com/stella/folio/commit/817c3cf4ed5399571dd88bbc0dbc3ce4313c5068), [`051dbd6`](https://github.com/stella/folio/commit/051dbd612dc6541df1725a29d7bfea8612bed056), [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697), [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697), [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697), [`3c137cd`](https://github.com/stella/folio/commit/3c137cd9e48be6b2199e504141fc8a4752558697), [`4f8eed6`](https://github.com/stella/folio/commit/4f8eed643be10173a177695b343e36e2d02d98cb)]:
  - @stll/docx-core@0.24.0

## 0.45.0

### Minor Changes

- [#884](https://github.com/stella/folio/pull/884) [`d457493`](https://github.com/stella/folio/commit/d457493202f6f93cb8a41d590b85aa12e4e5341b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a drawing's authored EMUs across the editor projection. Sizes, stroke
  widths, wrap insets and text-box margins were measured into pixels for the
  editor and converted back on save, and neither conversion is exact, so opening
  a document and saving it again moved every image, shape and text box off the
  numbers its author wrote. Each of the three nodes now carries the authored EMU
  beside the pixels it was projected into (`_docxAuthoredEmu` on `ImageAttrs`,
  `ShapeAttrs` and `TextBoxAttrs`) and writes it back while the pixel attribute
  still projects from it; a command that moves the pixels still reaches the
  document.

  Three defaults went with it, because each was written back as a value the
  document never had: the shape node's `outlineWidth` default of `1` gave an
  `a:ln` with no `@w` a width of 9525 EMU, the text-box node's margin defaults
  gave a text box with no `w:bodyPr` insets four authored ones, and a text box's
  authored inset of zero was dropped by a truthiness test. Every consumer already
  resolves an absent value against its own default.

- [#882](https://github.com/stella/folio/pull/882) [`9d05603`](https://github.com/stella/folio/commit/9d0560385dccd16738bc743b18d497b77016911a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound what parsing a package may allocate, in elements and attributes rather than only in bytes. The existing ceilings count inflated bytes, and bytes do not price a tree: measured on this repository's generators a parsed element retains about fifty bytes at its cheapest and a hundred and seventy at its densest, so `<w:r/>` costs seven bytes of markup and fifty of tree, and 128 MiB of markup inside every current bound buys some nineteen million elements and most of a gigabyte. Two holes made that reachable. The preflight ran on `word/document.xml`, `word/styles.xml` and `word/numbering.xml` by name, so headers, footers, footnotes, endnotes, comments and every other part were parsed into trees uncounted; and nothing bounded a package as a whole, so parts that were each unremarkable summed to the 250 MiB expansion ceiling, tens of millions of elements, and well past a gigabyte of tree.

  The preflight now runs on every XML part the unzip retains and on every part the server archive reader hands out as a string, in both cases against a package-wide budget, so a part added later is bounded by construction rather than by remembering to name it. `DocxUnzipLimits` gains `maxXmlElementsPerPart`, `maxXmlAttributesPerPart`, `maxXmlElementsPerPackage` and `maxXmlAttributesPerPackage`, and `DocxArchiveOptions` gains `xmlLimits`; all are host-configurable the same way the byte bounds are, and all are enforced by default. `XmlResourceLimitError` now carries the part path, the count reached where the scan stopped, and the bound it crossed. The defaults are drawn from the 5,314 readable packages in the public corpus, which reaches 602,212 elements and 639,110 attributes in a package: at 2,500,000 elements and 3,000,000 attributes a package they give better than four times the corpus maximum and reject nothing today's bounds accept, while capping a package at roughly 425 MB of tree. The byte bounds keep their values, and their doc comments now say what they do and do not cover.

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

### Patch Changes

- [#880](https://github.com/stella/folio/pull/880) [`966f842`](https://github.com/stella/folio/commit/966f8426c5308b4e8d6acf08087f2bee9f6d1ef0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve a SmartArt preview's drawing cache through the diagram's own `r:dm` id and the data part's `dsp:dataModelExt`, rather than by scanning the relationship map for the drawing type. The scan could serve one diagram another's drawing, and refused outright on a second match, so a document with two diagrams got a preview for neither.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Replay a preview-only drawing instead of rebuilding it. A `wpg:wgp` group is
  rendered to an SVG preview at parse time, and the model holds the render, not
  the group; once the preview's fingerprint no longer matched — after any edit,
  and on every rebuild path — the serializer regenerated DrawingML from the
  render, dropping the group's children, their relationships and the preview's
  own filename. Both classified raw-XML modes are now written back whatever the
  model says. A stale fingerprint still classifies the drawing `opaque`, which is
  how the lost edit is reported; it no longer licenses a replacement.

- [#880](https://github.com/stella/folio/pull/880) [`966f842`](https://github.com/stella/folio/commit/966f8426c5308b4e8d6acf08087f2bee9f6d1ef0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep content identity across the content-structure profile caps. A table past 128 blocks, or past the retained-text budget, had its exact signature, anchor texts and token counts blanked, and the fallback pairing read the absence as evidence: the only table in a document came back deleted and re-inserted when the document was compared with itself. The profile now carries a digest of every block at any size, so identical content pairs before any heuristic runs, and the heuristics it could not compute are modelled as `skipped-over-cap` rather than as empty.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `w:vMerge="restart"` on a cell whose merge has no continuation to span.
  The editor carries a merge origin as the cell's rowspan, which only exists once
  a continuation joins it, so every column whose merge closed at a rowspan of one
  lost its `w:vMerge` on save: a restart the table ends on, one a plain cell
  interrupts, one another restart supersedes. A merged cell losing its origin
  changes the table's visible structure.

- [#880](https://github.com/stella/folio/pull/880) [`966f842`](https://github.com/stella/folio/commit/966f8426c5308b4e8d6acf08087f2bee9f6d1ef0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Treat a list level neither document defines as the same absence rather than a difference. A `w:numPr` naming an abstract numbering with no such `w:ilvl` resolved to no level on either side, and the staging read that as a changed definition: the paragraph was remapped onto a freshly minted `numId`, so a document compared with itself reported a numbering change and any direct `w:ind` the new numbering displaced could no longer be moved back.

- [#880](https://github.com/stella/folio/pull/880) [`966f842`](https://github.com/stella/folio/commit/966f8426c5308b4e8d6acf08087f2bee9f6d1ef0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `w:hyperlink/@w:tgtFrame` as the document authored it. Any frame name outside `_blank`, `_self`, `_parent` and `_top` was mapped to `_blank` at parse, so a saved file no longer said what the source said. The allow-list clamp now lives where a DOM anchor or a navigation is produced, in one owner (`anchorTargetAttrs`) that every rendered document, editor popover and `window.open` in core, React and Vue goes through, so the `target` and the `rel` that must accompany it are decided in a single place.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop giving a shape outline a dash pattern it never stated. An `a:ln` with no
  `a:prstDash` came back through the editor as `style: "solid"`, which is the
  shape's own decision rather than the absence the source had, so a later change
  to what an unstated outline renders as could no longer reach it. The shape
  node's `outlineStyle` now defaults to absent, and the renderer already draws an
  unstated outline solid.

- [#884](https://github.com/stella/folio/pull/884) [`d457493`](https://github.com/stella/folio/commit/d457493202f6f93cb8a41d590b85aa12e4e5341b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read `name=""` on a drawing object as no name. `@name` is schema-required on
  `CT_NonVisualDrawingProps`, so a shape, text box or picture the model never
  named still writes one, and the reader took that empty string back as authored
  content: `absent → save → parse` landed on `""` instead of absent, and the next
  save carried it. Nothing downstream can tell the two apart, so the reader now
  maps the one value the writer mints back to absence. `@descr` and `@title` are
  optional and written only when authored, so `""` in either stays a string
  someone wrote.

- [#882](https://github.com/stella/folio/pull/882) [`9d05603`](https://github.com/stella/folio/commit/9d0560385dccd16738bc743b18d497b77016911a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Encode the SmartArt preview raster without walking it a byte at a time. Every diagram drawing rasterises a bounded megapixel placeholder at parse time, but the PNG's CRC and Adler checksums iterated it with `for (const byte of bytes)`, and the data URI was built by spreading thirty-two thousand arguments per chunk into `String.fromCodePoint`. A profile of a four-diagram package put three quarters of the whole parse in the array-iterator protocol. The checksums now run over indices with a CRC table and a blocked Adler accumulator, the zlib stream is finished in its own buffer rather than copied into a second one the size of the raster, and base64 is written once into an ASCII array. The bytes produced are unchanged.

- [#877](https://github.com/stella/folio/pull/877) [`601e5a7`](https://github.com/stella/folio/commit/601e5a79e4ba4f573930a89d22ebce7eaf3adb01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every comment's author, body, date, resolved state and reply link with its own `w:id` across a save. `word/comments.xml` was written with the top-level comments first and the replies after, so a document whose comments.xml interleaves a reply with a later thread root came back from the next parse in a different `comments[]` order than it went in, and anything reading that array by position saw one comment's text and author under another's place. Comments are now written in the model's order, both comment parts are planned once from one ordered list of ids, and a duplicate or missing `w14:paraId` resolves to the same comment on parse and on save.

- [#881](https://github.com/stella/folio/pull/881) [`05044c5`](https://github.com/stella/folio/commit/05044c53a4b02600a661c834cb65b09d5f31a56f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read and write a decorative image as the extension Word writes, and keep `hidden` a separate fact. `Image.decorative` was read from a `@decorative` attribute `CT_NonVisualDrawingProps` does not have (no file in the public corpus writes one), and written back as `hidden="1"`, which says the drawing is not displayed — so a decorative image became a hidden one, and re-parsed as neither. It now round-trips through `wp:docPr`'s `{C183D7F6-B498-43B3-948B-1728B52AA6E4}` extension, `Image.hidden` carries `@hidden` on its own and is written identically for inline and anchored drawings, and `Image.docPrExtensions` keeps the other `a:ext` entries of the same list verbatim and in order rather than dropping them. All three survive the editor round trip.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Save `w:keepNext`, `w:keepLines` and `<w:specVanish/>` when a command sets
  them. All three were classified `original-only` in the paragraph write-back
  map: an imported value survived through `_originalFormatting`, and a value set
  on the paragraph node had no save path at all, so a paragraph with no `w:pPr`
  of its own lost it silently. They join `widowControl` as `style-resolved-attr`,
  so a commanded value is written and a value that only echoes the paragraph's
  style still is not.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Never write a content-control selection nobody made. A block-level dropdown or
  combo box whose `w:sdtPr` was not replayed verbatim — every control an editor
  command built, and every control on the rebuild path — had its `@w:lastValue`
  recovered from the body's display text, so a control still showing its
  placeholder was saved as selected, and a displayText shared by two list items
  selected the first of them. `properties.dropdownLastValue` is now the only
  record of a selection; the schema's empty-string default keeps "never
  selected", "cleared" and "selected" distinguishable.

- [#882](https://github.com/stella/folio/pull/882) [`9d05603`](https://github.com/stella/folio/commit/9d0560385dccd16738bc743b18d497b77016911a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound the SmartArt preview a parse generates. Only the VML preview was charged against a package-wide budget, and it was recognized by constants in a different file from the ones its producer wrote, so a rename would have silently stopped the charge. Both previews now come from one table that the producer builds from and the budget matches against, and each kind carries its own per-package allowance. The SmartArt cap is set above the public corpus maximum (51.3 MB of preview data URL, from a package under a megabyte), so no corpus file loses a preview that it keeps today.

- [#882](https://github.com/stella/folio/pull/882) [`9d05603`](https://github.com/stella/folio/commit/9d0560385dccd16738bc743b18d497b77016911a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Encode bytes to base64 through one owner. Four call sites each built `btoa`'s binary string their own way, and the markdown renderer's way was wrong: `TextDecoder("latin1")` is the windows-1252 decoder by specification, so any byte in 0x80-0x9F produced a character `btoa` rejects and registering an ordinary image threw `InvalidCharacterError` in browsers, where no `Buffer` fallback hides it. `utils/base64` now encodes bytes directly, using the runtime's `Uint8Array.prototype.toBase64` where there is one, and a lint rule keeps `btoa` out of package source.

- [#881](https://github.com/stella/folio/pull/881) [`05044c5`](https://github.com/stella/folio/commit/05044c53a4b02600a661c834cb65b09d5f31a56f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give XML escaping one owner, and make its output always well-formed. `@stll/docx-core` now exports `escapeXmlText` and `escapeXmlAttribute`: six hand-rolled escapers disagreed about the characters that matter, so a value could leave folio as markup Word refuses to open, or come back changed. Both functions drop the characters XML 1.0 §2.2 forbids (the C0 controls outside tab/LF/CR, U+FFFE, U+FFFF, unpaired surrogates), which cannot be escaped into a document either. The attribute form writes tab, LF and CR as character references, because §3.3.3 has every conformant reader flatten a literal one to a space; the text form does the same for CR, which §2.11 would otherwise rewrite to LF. `sanitizeXmlCharacters` applies the same rule at an input boundary, where the value can still be reported.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write page-border art relationship ids in the relationships namespace. `r:id`
  on every `w:pgBorders` side, plus `r:topLeft` / `r:topRight` on the top and
  `r:bottomLeft` / `r:bottomRight` on the bottom, were written as `w:id`,
  `w:topLeft` and so on. Those are different attributes: Word discarded them and
  the border art with them, and folio's own prefix-tolerant reader hid it by
  reading its own output back. An attribute-less `<w:docGrid/>` is written back
  too, rather than dropped for having nothing to say. Nine pairs leave the
  container survival baseline.

- [#882](https://github.com/stella/folio/pull/882) [`9d05603`](https://github.com/stella/folio/commit/9d0560385dccd16738bc743b18d497b77016911a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Serialize captured XML directly instead of through a parallel node tree. `elementToXml` built a second copy of every subtree in fast-xml-parser's builder format before writing it, and captures nest, so the same bytes were copied at every level on the way out. It now appends into one shared buffer. The output is unchanged: a differential property test compares it against the builder it replaced over generated trees, and the two agree on every one of the 194,413 elements in a 120-package corpus sample.

- [#879](https://github.com/stella/folio/pull/879) [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry a shape outline's `a:ln@join` through the ProseMirror model. The parser
  read it and the serializer wrote it, but the shape node had nowhere to put it,
  so a mitred or bevelled outline came back rounded after any edit.

- [#880](https://github.com/stella/folio/pull/880) [`966f842`](https://github.com/stella/folio/commit/966f8426c5308b4e8d6acf08087f2bee9f6d1ef0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compare inline-atom topology on document facts alone. A text box projects a `textBoxAnchor` whose id is minted per conversion and salted with a random nonce, and the key both sides were compared on carried that id verbatim, so any paragraph holding a text box beside a field, image or page break failed to align — a document differed from itself. The key now drops attributes a conversion mints for itself, and `offsetAt` no longer claims a refusal it could not return.
- Updated dependencies [[`05044c5`](https://github.com/stella/folio/commit/05044c53a4b02600a661c834cb65b09d5f31a56f), [`a56ab6a`](https://github.com/stella/folio/commit/a56ab6a0dd29cb0b5810813a4bc36eadec0735a3), [`05044c5`](https://github.com/stella/folio/commit/05044c53a4b02600a661c834cb65b09d5f31a56f)]:
  - @stll/docx-core@0.23.0

## 0.44.0

### Minor Changes

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `w:bdo` and `w:dir`, the Unicode bidirectional controls. `w:dir` is an embedding and `w:bdo` an override, and folio discarded both on save — in a right-to-left document that is the difference between a readable line and a scrambled one, because an override is what makes a Latin word inside it read backwards. They are now a `BidiWrapper` member of `ParagraphContent`, a transparent inline container that nests, holds anything paragraph content holds, and that every paragraph walk reads straight through. The editor projection flattens the wrapper for now, so a document edited in the editor still loses the direction; the save path keeps it.

- [#866](https://github.com/stella/folio/pull/866) [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672) Thanks [@jan-kubica](https://github.com/jan-kubica)! - No document folio authors references a style it does not define. A comment or note reference mark carried `w:rStyle` pointing at `CommentReference`/`FootnoteReference`/`EndnoteReference` whether or not the package declared them, so the mark lost its superscript; the generic style set stopped at `Heading4` while the report builder accepts six levels and applies `TableGrid`; and `generateTOC` wrote `TOCHeading`/`TOC1` regardless of what the open document calls its TOC styles. The table of contents now takes its styles from the document through the built-in classifier, and writes no style id rather than a dangling one. A property test holds every authoring path to this.

  **Breaking for direct callers of the TOC command:** `generateTOC` is now `generateTableOfContents({ title })` and `insertTableOfContentsInView(view, { title })`. The title was a hardcoded English "Table of Contents", which is wrong in every document that is not in English; this layer has no locale, so the host supplies the string.

- [#868](https://github.com/stella/folio/pull/868) [`170e3ca`](https://github.com/stella/folio/commit/170e3cad6254f7c372c9dff131584df70d3f85be) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `Document.parseWarnings` reports what the parse boundary normalised, as data: a stable `code` from `PARSE_WARNING_CODES`, the part it happened in, the best position that part can name, and the value folio declined to read. `Document.warnings` is unchanged in shape and is now rendered from that list by one formatter, so the prose and the data cannot disagree. Normalisations that were silent now report: a `w:type` outside `ST_HdrFtr`, a repeated footnote or endnote id, a value outside `ST_OnOff` in either shape, a border with no `w:val`, a `w:comment` with no readable `w:id` (previously read as id 0, which manufactured a duplicate), and a hyperlink naming a relationship its part never defined. Retained warnings are capped per code, with the remainder counted.

### Patch Changes

- [#871](https://github.com/stella/folio/pull/871) [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Mint revision ids out of the whole annotation space. `w:id` on a comment, a bookmark, a protected range and a tracked change comes from one counter, but the save-time deduplication reserved only revision ids and handed out the lowest free integer, so renumbering a duplicated `w:ins` in a commented document could land on a live comment's id. The pass now reserves every annotation id (never claiming one, since a comment id legitimately repeats), and both editor allocators seed above the maximum of the whole space rather than their own kind.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Record the `w:tblLook` that applying a table style asks for. `applyTableStyle` painted the first row, the last row and the bands the style names but left the table's own look untouched, so a save kept whatever the document arrived with and Word resolved the conditional formatting differently from what was on screen. The command now states each of the six regions explicitly and leaves `w:val` as the author wrote it: the attribute form is what a reader resolves first, and re-encoding the bitmask would rewrite bits folio does not model.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a bare OMML element that sits in paragraph content. Every group that admits `m:oMath` also admits `m:EG_OMathMathElements`, so `<w:ins><m:f/></w:ins>` is a tracked insertion of a fraction with no `m:oMath` around it. The parser recognised only `m:oMath` and `m:oMathPara` and let the rest fall off the end of its `switch`, so a tracked wrapper reached disk with its content gone — a reviewer accepting an edit that is no longer there. Bare equations now travel as the markup they arrived as, exactly like the ones that have a wrapper. 76 of the survival census's pairs move from lost to kept.

- [#871](https://github.com/stella/folio/pull/871) [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the range markers that stand between two blocks. `w:body`, `w:tc`, a header and an SDT's content all admit `w:permStart`, `w:customXml*Range*` and a comment or move range beside their paragraphs, and every block container dropped them on save: a protected range lost its `w:permStart` and the saved file came back unprotected. The markers are now captured verbatim and replayed where they stood, the way an SDT's sibling markers already were. They do not yet survive the editor round trip, which needs a zero-width node rather than a block attribute.

- [#867](https://github.com/stella/folio/pull/867) [`2538df2`](https://github.com/stella/folio/commit/2538df2dc8c6957ca520c9828f776e3bfb216f41) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `nil` and `none` distinct wherever a `CT_Border` is read or written. They are two members of `ST_Border`, not synonyms, and the build-from-scratch serializer rewrote `none` as `nil`. The four `parseBorderSpec` copies also collapse into one reader, so a border element with no `w:val`, an explicit `w:shadow="0"` and the page-border art relationship ids are now read the same way on the paragraph, style, table, cell and page tiers. The light grid a generated table gets when it declares no borders is unchanged, but is now named as the authoring default it is.

- [#866](https://github.com/stella/folio/pull/866) [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Every style table folio authors names its built-ins with the `w:name` Word itself writes, from one constant map in `docx/builtInStyles.ts`. The generic style set wrote `Heading 1`…`Heading 4` where Word writes `heading 1`, and the Stella set wrote `Footnote Text`, `Footnote Reference`, `Endnote Text`, `Endnote Reference` and `Footer` where Word writes those lowercase. Word's own casing is not uniform, so the map records each spelling separately. The TOC-entry style matcher now uses the same name normaliser as the classifier, accepting `toc 1`, `TOC 1` and `TOC  1` alike.

- [#871](https://github.com/stella/folio/pull/871) [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Make saving a threaded comment set a fixed point. `word/comments.xml` is written with the top-level comments first and the replies after, so the next parse returns the comments in that order, but `word/commentsExtended.xml` was built by walking the model's own order: the first save wrote an entry order the second save could not reproduce. Both parts, and the paraId minting they key on, now walk one order.

- [#872](https://github.com/stella/folio/pull/872) [`84c1650`](https://github.com/stella/folio/commit/84c1650475771bc1eeb18ee64ad5605e50089469) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Comparing a document with itself reports no change in three cases where it used to refuse or invent one. An inline atom the comparison cannot detach from its package — an image with no embedded media, a drawing that is a chart rather than a picture — is now compared as part of the block's topology instead of making the whole story unalignable. A package with no `word/styles.xml` compares against another that has none: two packages with no style definitions share one formatting context, so there is nothing to isolate and nothing to read a definition from. And two structurally identical tables no longer cancel each other's pairing: each shares every signature with the other by construction, and evidence that also exists locally no longer counts as evidence that the counterpart lies elsewhere.

- [#873](https://github.com/stella/folio/pull/873) [`48715e3`](https://github.com/stella/folio/commit/48715e31048aabd8ad2c27d5d647aefcb20a6a53) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Derive `xml:space="preserve"` from the text instead of storing it. `TextContent.preserveSpace` is gone: whether `<w:t>` needs the attribute is a pure function of its text, and a stored copy of a derived fact only drifts — the editor lost it whenever two adjacent runs merged, because ProseMirror has nowhere to carry it. Both serializers now call the same `requiresXmlSpacePreserve`, which `@stll/docx-core` exports.

- [#872](https://github.com/stella/folio/pull/872) [`84c1650`](https://github.com/stella/folio/commit/84c1650475771bc1eeb18ee64ad5605e50089469) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A drawing with no picture relationship keeps the markup it arrived with. A `w:drawing` whose graphic is a chart or an OLE frame, or which carries no `a:graphic` at all, has no `a:blip` and so no relationship id; `Image.rId` is now absent in that case rather than an empty string, and a save writes the anchor back as authored instead of rebuilding it into a picture bound to whichever relationship the part happens to list first. Relationship ids resolve through one typed resolver that distinguishes a resolved id from an absent and a dangling one, and an image reference that names a non-image relationship no longer resolves to that part.

- [#873](https://github.com/stella/folio/pull/873) [`48715e3`](https://github.com/stella/folio/commit/48715e31048aabd8ad2c27d5d647aefcb20a6a53) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep an optional paragraph property that a style supplies out of the paragraph's own `w:pPr`. The editor holds the effective value, and the save path read it as authored, so `w:bidi`, `w:widowControl`, `w:suppressAutoHyphens`, `w:kinsoku`, `w:overflowPunct`, `w:snapToGrid`, `w:pageBreakBefore`, `w:contextualSpacing`, `w:ind`, `w:pBdr`, `w:shd`, `w:tabs` and `w:outlineLvl` came back as direct formatting on any document whose styles define them: the paragraph stopped following its style, and an unset tri-state became an explicit value. A table with no `w:tblBorders` no longer gets one synthesised from its first bordered cell for the same reason. A run carrying an embedded object, picture or shape keeps its `w:rPr`.

- [#869](https://github.com/stella/folio/pull/869) [`0620ed7`](https://github.com/stella/folio/commit/0620ed7f95babbcc65f2c3a6746da8c7772ed500) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `ExhaustiveFields<Source, Classified>`, the compile-time gate that turns a model field added without a decision into a build failure, now has one owner and is exported from `@stll/docx-core/model`. The paragraph, text, and border serializers each carried a verbatim copy.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read every model content union by exhaustion. The conversion, serializer and markdown modules narrowed `ParagraphContent`, `BlockContent`, a hyperlink's children and the table structural-change union by chains of `else if`, so a member added to any of them reached whichever branch happened to be last and the build still passed. Each is now a `switch` with a `never` default, and a lint rule fails the next chain written in those directories.

  Three bugs the conversion surfaced: a bookmark pair inside `w:bdo`/`w:dir` was not recognised as a pair, so the start fell back to the legacy paragraph attribute and the end was dropped; a text-box anchor inside one was neither resolved nor removed and reached the saved package; and a footnote or comment holding an equation rendered as markdown without it. Two unused plain-text flatteners that disagreed with the one owner are gone.

- [#866](https://github.com/stella/folio/pull/866) [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Headings are recognised from the paragraph's effective `w:outlineLvl` and the style's built-in `w:name` rather than from an English style id. A localized Word writes `Nadpis1`, `berschrift2`, `Titre3`, `Nagwek4` or `Cmsor5` for the same built-in heading, so id matching found English output and nothing else: those paragraphs were missing from the outline sidebar and generated tables of contents, arrived at the AI snapshot as plain paragraphs, and exported to Markdown without `#`. One classifier (`docx/builtInStyles.ts`) now answers for every consumer, with `w:outlineLvl` 9 meaning body text rather than a tenth level, and the bilingual builder's language-list regex is gone. Documents folio creates carry outline levels on their heading styles, and define the `Quote` style a Markdown blockquote compiles to.

- [#867](https://github.com/stella/folio/pull/867) [`2538df2`](https://github.com/stella/folio/commit/2538df2dc8c6957ca520c9828f776e3bfb216f41) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a `w:val` outside `ST_OnOff` as absent in both shapes of the type. The element shape (`<w:b w:val="yes"/>`) used to read anything it could not parse as `true` while the attribute shape read the same malformed value as nothing, so one type answered the same input two ways. A census of 5,316 public documents from 289 producers found no producer writing a non-standard spelling systematically, so no tolerance beyond `1`, `0`, `true`, `false`, `on` and `off` is added.

- [#866](https://github.com/stella/folio/pull/866) [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A heading pasted at a level the open document defines no style for lands on the deepest heading style it does define, instead of keeping a style id that resolves to nothing.

- [#866](https://github.com/stella/folio/pull/866) [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A heading pasted from another application is pointed at the open document's own heading style instead of the English built-in id the schema's paste rule can produce. In a document whose heading styles are localized, or one whose style set stops at level four, that id resolved to nothing and the pasted heading lost the document's heading formatting. A paragraph is retargeted only when the document defines no style under the id it carries and does define one at that outline level.

- [#871](https://github.com/stella/folio/pull/871) [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every attribute a range marker arrived with. `w:moveFromRangeStart` and `w:moveToRangeStart` lost `w:author` and `w:date`, which their schema type requires, so a saved document was markup Word repaired; `w:displacedByCustomXml` was lost on every bookmark, comment range and move range. The markers now model the schema's own `CT_MarkupRange` / `CT_Bookmark` / `CT_MoveBookmark` chain and share one reader and one writer.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalise an unbalanced range marker wherever it sits. The model validator walks a paragraph's whole inline tree; the two parse-boundary normalisers walked `paragraph.content` alone, so a `w:commentRangeStart` or a move-range marker inside `w:ins`, `w:hyperlink`, `w:sdt`, `w:bdo` or `w:dir` was judged but never normalised and made `parseDocx` throw on a document Word opens. Both now read the tree through one exhaustive traversal. The per-kind policy is unchanged: a comment range's unmatched half becomes a point `w:commentReference`, a move range's is dropped.

- [#872](https://github.com/stella/folio/pull/872) [`84c1650`](https://github.com/stella/folio/commit/84c1650475771bc1eeb18ee64ad5605e50089469) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Every relationship-id lookup goes through the typed resolver. A hyperlink whose `r:id` names a non-hyperlink relationship no longer resolves to that part's path; a `w:headerReference` or `w:footerReference` with no `r:id` is dropped at parse instead of entering the model as an empty id and being written back as the schema-invalid `r:id=""`; a header or footer part with no `.rels` of its own resolves against nothing rather than against the document's relationships, so a part-local `rId1` can no longer name a body part; and a VML `v:imagedata` or grouped picture with no id records absence rather than an empty string. A body reference whose id names nothing is reported in the parse warnings, so a dangling reference is distinguishable from one the author never wrote.

- [#867](https://github.com/stella/folio/pull/867) [`2538df2`](https://github.com/stella/folio/commit/2538df2dc8c6957ca520c9828f776e3bfb216f41) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read each OOXML reserved value through a single owner, so no two call sites can disagree about it. `w:u w:val="none"` now cancels an inherited underline everywhere instead of painting a solid one when text is typed; `wp:anchor behindDoc="true"` puts a shape or text box behind the text, as `behindDoc="1"` already did; every `ST_OnOff` attribute accepts `on` and `off`, so `w:beforeAutospacing="on"` is no longer read as its opposite and saved back inverted; `w:shd w:fill="auto"` survives a save on paragraphs, styles and table cells, not only on runs; and `w:tblW`/`w:tcW` with `w:type="auto"` autofit instead of being pinned to the meaningless width Word leaves in `w:w`.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Open a `w:sym` that names only one of its two optional attributes. `CT_Sym` declares `w:font` and `w:char` optional and Word renders a `<w:sym w:char="F0B7"/>` by falling back to the run's font; folio refused the document twice over, once in the model validator and once in the ProseMirror projection, so a file Word opens did not open at all. Both checks now accept an absent attribute and still reject a malformed character, and the serializer writes an absent attribute back as absent instead of inventing `w:font=""`.

- [#868](https://github.com/stella/folio/pull/868) [`170e3ca`](https://github.com/stella/folio/commit/170e3cad6254f7c372c9dff131584df70d3f85be) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A `DocumentStyleSet` is normalised where it enters folio rather than trusted. A set persisted as JSON before folio learned to repair one could still carry a style numbering it never defines, two styles under a single id, or an initial paragraph style it does not contain, and `createEmptyDocument` would panic on the first of those. It now repairs all three through the owners the parser and the extractor already use, leaves the caller's value untouched, and reports each repair on the resulting document. `DOCUMENT_STYLE_SET_VERSION` is unchanged: the shape did not move.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a table width spelled as a percentage as one. `w:type` is optional on `CT_TblWidth` and the schema gives it no default, so folio read a missing one as `dxa` and `<w:tblW w:w="50%"/>` became 50 twips, then saved that way: a table half the page wide came back a hairline. `w:w` is `ST_MeasurementOrPercent`, so a `%` spelling is a percentage whatever `w:type` says; which slots admit it and what unit their number counts in now comes from the generated slot table the verbatim capture already uses, so the reader and the capture cannot disagree about one width.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep every `w:tblLook` the author wrote. `CT_TblLook` states which conditional formats a table takes from its table style twice: as the legacy `w:val` bitmask, which `TableLook` had no field for, and as six `ST_OnOff` attributes, which the serializer wrote only when true. A rebuild therefore turned `w:val="04A0" w:firstRow="1" w:lastRow="0" w:noHBand="0"` into `w:firstRow="1"`, and the two are different documents: an absent flag falls back to `w:val`'s bit, an explicit `0` overrides it. `TableLook` gains `val` and each flag is now tri-state. Two readers also disagreed about precedence — the table one OR-ed `w:val`'s bits over an explicit `0`, so a table that switched its header row off got one anyway; `styleParser` now calls the table parser, and `resolveTableLook` is the single place a flag resolves to an answer.

- [#875](https://github.com/stella/folio/pull/875) [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `w:tblGridChange` and `w:numberingChange` when the container they sit in is rebuilt. Both are tracked-change records — the grid a reviewer replaced when resizing a column, and the numbering a reviewer replaced when changing a list — and nothing in the editable model derives either, so a save that rebuilt the container dropped the revision and the document then read as though the change had always been there. They now travel as their own capture slots (`TableFormatting.gridChangeXml`, `ParagraphFormatting.numberingChangeXml`) and are written back on both the replay and the rebuild path. A captured `w:pPr` carrying a `w:numberingChange` is no longer refused for replay either; refusing it used to force the rebuild that could not write it.
- Updated dependencies [[`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410), [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016), [`2538df2`](https://github.com/stella/folio/commit/2538df2dc8c6957ca520c9828f776e3bfb216f41), [`48715e3`](https://github.com/stella/folio/commit/48715e31048aabd8ad2c27d5d647aefcb20a6a53), [`84c1650`](https://github.com/stella/folio/commit/84c1650475771bc1eeb18ee64ad5605e50089469), [`0620ed7`](https://github.com/stella/folio/commit/0620ed7f95babbcc65f2c3a6746da8c7772ed500), [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672), [`c0e1b75`](https://github.com/stella/folio/commit/c0e1b75de80787791d4119f170fdbf2d54593672), [`72a12b0`](https://github.com/stella/folio/commit/72a12b05e3beccdfeb1b977256c9b5549dc38016), [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410), [`170e3ca`](https://github.com/stella/folio/commit/170e3cad6254f7c372c9dff131584df70d3f85be), [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410), [`8a90370`](https://github.com/stella/folio/commit/8a90370e19275d4ce6f4adf1e9f48c13fb556410)]:
  - @stll/docx-core@0.22.0

## 0.43.0

### Minor Changes

- [#858](https://github.com/stella/folio/pull/858) [`d09bc26`](https://github.com/stella/folio/commit/d09bc26e0997b1eee0492dfda65bb6ab5e456f4a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Turn a list marker typed at the start of a paragraph into a list item, the way Word autoformats as you type: `- ` and `* ` start a bulleted list, `1. ` a numbered one. The rule runs the command the toolbar list buttons run, so an autoformatted list carries the same numbering properties and saves the same way, and Backspace right after the conversion puts the typed marker back. A marker typed anywhere but the start of a plain paragraph stays text.

- [#856](https://github.com/stella/folio/pull/856) [`ee366fb`](https://github.com/stella/folio/commit/ee366fb1ce350813568dc54d285d89d19ab13ed3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The template fill preview acts on conditional blocks a host has ruled on: `TemplatePreviewValues.conditions` maps an `{% if %}` expression, or the bare field path its filter chain hangs off, to whether the block applies; `false` drops the span from the opener through its `{% endif %}`, `plain` mode also drops the tag lines of a block that does apply, and the paged layout drops the blocks those spans swallow whole so the pages paginate without them.

### Patch Changes

- [#865](https://github.com/stella/folio/pull/865) [`cb31905`](https://github.com/stella/folio/commit/cb319050517c7255b38f8c6c9b75cc5579870847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The default paragraph style is resolved the way ECMA-376 17.7.4.17 does rather than by assuming the style id `Normal`: the paragraph style flagged `w:default="1"` wins, the last one where several are flagged, then a style carrying the built-in `w:name`, then the built-in id. A localized or generated package that names its default `Standard`, `Normln` or `style0` now resolves it, and one that declares no default gets a minted default instead of a failed extraction.

- [#865](https://github.com/stella/folio/pull/865) [`cb31905`](https://github.com/stella/folio/commit/cb319050517c7255b38f8c6c9b75cc5579870847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A `word/comments.xml` that defines two comments under one `w:id` now parses. The body addresses a comment by that id, so Word resolves every marker naming it to the first definition; folio keeps the first, drops the later ones no marker can address, and reports it as a parse warning. Footnotes and endnotes keep the first definition of a repeated id too, which their id index already did.

- [#865](https://github.com/stella/folio/pull/865) [`cb31905`](https://github.com/stella/folio/commit/cb319050517c7255b38f8c6c9b75cc5579870847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A full repack no longer refuses a document whose `w:headerReference` or `w:footerReference` states a `w:type` outside `ST_HdrFtr`, such as the `odd` some producers write for the default header. The reference-loss guard reads both sides of its comparison through the same parser, so the value the parse boundary normalised is no longer mistaken for a dropped reference. What the guard refuses is unchanged.

- [#860](https://github.com/stella/folio/pull/860) [`f5919f1`](https://github.com/stella/folio/commit/f5919f1f2a3df2cadafc1c82bd61f453b92b172b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A template preview value carrying newlines now breaks its lines inside the marker's paragraph instead of painting them on top of the content below it: each newline becomes a line break run, the way `w:br` does, so the paragraph measures the height its value needs.

- [#862](https://github.com/stella/folio/pull/862) [`7c18a6e`](https://github.com/stella/folio/commit/7c18a6e38b31162782ee154ebb7a6de52297125d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `w:numId w:val="0"` on a paragraph style is ECMA-376's "no numbering" sentinel, not a reference: it switches off the numbering the style would inherit through `w:basedOn`. Saving a document whose styles carry it (a TOC heading based on a numbered heading, for example) no longer fails with "Style references missing numbering definition 0", and the sentinel is written back unchanged. Every numbering lookup now reads the sentinel through one shared predicate.

- [#865](https://github.com/stella/folio/pull/865) [`cb31905`](https://github.com/stella/folio/commit/cb319050517c7255b38f8c6c9b75cc5579870847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A DOCX with no `word/styles.xml` no longer fails style-set extraction. The package is valid and Word opens it on its built-in defaults, so the extracted set carries a minted default paragraph style. Where the source declared no `w:docDefaults`, that style carries the built-in Normal formatting too, because a set that names a default paragraph style is no longer a package a consumer applies its own built-in to.

- [#865](https://github.com/stella/folio/pull/865) [`cb31905`](https://github.com/stella/folio/commit/cb319050517c7255b38f8c6c9b75cc5579870847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Two page-break shapes Word writes routinely now convert. A page break opening a table cell projects as the row's break whatever else the cell holds, where before the cell had to contain that one paragraph alone. A paragraph carrying both a text-box anchor and a page break converts when the anchor precedes the break, which is the arrangement layout already projects faithfully; only an anchor that follows the break is still refused.

- [#863](https://github.com/stella/folio/pull/863) [`0c6a3f8`](https://github.com/stella/folio/commit/0c6a3f87ce97e2f62392b9b916e4df4d33fcbdc5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A `w:numPr` naming numbering the file never defines no longer fails a save. Parsing and style-set extraction rewrite such a reference to the `w:numId w:val="0"` sentinel on both the paragraph and the paragraph style, with a warning naming the style, so a style whose numbering is missing stays unnumbered instead of inheriting its `w:basedOn` parent's list; deleting the `w:numPr`, which the paragraph tier did before, numbered a paragraph its source showed unnumbered. Building a package now checks the numbering a style names, not every `w:num` the source carried.

- [#857](https://github.com/stella/folio/pull/857) [`1f3a9b6`](https://github.com/stella/folio/commit/1f3a9b654c8eb02416aaf82a96334acdf0fc46c3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Start each vertical caret step from where the caret actually is. ArrowUp and ArrowDown remembered the visual line and column of the previous step and only discarded them on a key the editor view handled, so a caret moved by anything else (a click, a find result, an agent edit) kept stepping from the line it had left: the next ArrowDown skipped a visual line and snapped back to the earlier column. The remembered step is now bound to the position it settled on, which keeps it where it is needed — a soft-wrap boundary belongs to both the line it ends and the line it starts — and re-resolves it everywhere else. A caret after a line-edge space also measures its column through the painted caret geometry, so a vertical step keeps the column the caret is drawn at instead of the one before the space.

## 0.42.0

### Minor Changes

- [#854](https://github.com/stella/folio/pull/854) [`9e15419`](https://github.com/stella/folio/commit/9e154197039506f9617dc58d01e566e3ce9ae5eb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Document properties of a newly created document are configurable.

## 0.41.0

### Minor Changes

- [#848](https://github.com/stella/folio/pull/848) [`66f0734`](https://github.com/stella/folio/commit/66f0734a93dc0dc78e41dacac4a7414dc9176327) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Classify a drawing as `native`, `replayable` or `opaque` through the predicate the run serializer already uses, so a document is no longer opened read-only because a header carries a logo; only `opaque` content blocks editing, and `DocxCompatibility` gains a `drawings` list at `schemaVersion: 2`. Regenerating a picture now round-trips `a:graphicFrameLocks` and `wp:effectExtent`, and a rasterized shape group is marked `previewOnly` so the editor declines to resize it rather than replacing the group with one child picture. Shape drawings Folio cannot model — unmodeled effects and 3-D, `wpg:wgp` groups without a preview, a `w:pict` with no resolvable image, an `mc:AlternateContent` whose every branch declines — are preserved verbatim instead of dropped. Field results are no longer missing from the AI-facing block text, so a paragraph carrying a cross-reference reads as the text Word shows.

### Patch Changes

- [#850](https://github.com/stella/folio/pull/850) [`46065eb`](https://github.com/stella/folio/commit/46065ebe7176383de5fc6208fa161109cf00fb37) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a field in note, header and footer text as its stored result rather than the text the editor paints for it. A field with no result used to contribute a synthesized placeholder, and a DATE field the current date, so the same document produced different text on different days and anything hashing or comparing that text was unstable. Every other inline atom contributes exactly what it did before.

- [#852](https://github.com/stella/folio/pull/852) [`6f9af7e`](https://github.com/stella/folio/commit/6f9af7ec53d318527629afe3e067e8916e0edcfc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Read a story's text through one projection whether or not it is open for editing. A loaded story used to run a separate walk that glued words together across a tab or a hard break, showed tracked-deleted text, and ran table cells and paragraphs into each other, so the same footnote read differently before and after it was loaded. Header and footer text came from a third walk that saw only plain runs, silently dropping fields, hyperlinks, tabs and breaks. All of them now read the document model's own walk, and a result-less PAGE field reports what the document holds rather than the placeholder a save writes.

- [#853](https://github.com/stella/folio/pull/853) [`e7b2294`](https://github.com/stella/folio/commit/e7b22941dd368750e693b73470fad561554e8346) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Name the element type of the operation contract's alignment and line-spacing value lists. Spreading a const tuple widened them to an array of the union, and the declaration emitter wrote that union out member by member in an order that changed between builds, so the emitted declarations were not reproducible.

- [#851](https://github.com/stella/folio/pull/851) [`07ac618`](https://github.com/stella/folio/commit/07ac618ce3ff5f83e294bc330a52506845a7ef46) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Name the element type of two derived constants so the emitted declarations stop varying between builds. Their inferred type was a union the declaration emitter wrote out member by member in an order that changed from build to build, which made `dist/**/*.d.ts` non-reproducible.
- Updated dependencies [[`66f0734`](https://github.com/stella/folio/commit/66f0734a93dc0dc78e41dacac4a7414dc9176327)]:
  - @stll/docx-core@0.21.0

## 0.40.0

### Minor Changes

- [#846](https://github.com/stella/folio/pull/846) [`8fc69b7`](https://github.com/stella/folio/commit/8fc69b78abe4db6bd6078fa9b4ea280137196a44) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep `listRendering.levelStarts` through a Document → ProseMirror → Document rebuild so custom list starts render without a DOCX round-trip, and add `formattingScope: "allParagraphs"` to block insertions so a multiline `text` can produce several list items. `ListRendering.levelStarts`, `DocumentSettings.mirrorMargins`, and the header/footer verbatim capture fields are now declared on the model types instead of attached through local intersections.

### Patch Changes

- Updated dependencies [[`8fc69b7`](https://github.com/stella/folio/commit/8fc69b78abe4db6bd6078fa9b4ea280137196a44)]:
  - @stll/docx-core@0.20.2

## 0.39.1

### Patch Changes

- [#843](https://github.com/stella/folio/pull/843) [`832f46a`](https://github.com/stella/folio/commit/832f46ae1330ba2cb73f61ca041ee8b01b323454) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route committed browser text input through shared model transactions to preserve revision boundaries in both adapters.

## 0.39.0

### Minor Changes

- [#837](https://github.com/stella/folio/pull/837) [`4f314d2`](https://github.com/stella/folio/commit/4f314d2b8e7b0ae38e6fe276cbd8cdf1adc8feee) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor centered and bottom-aligned section content in paginated views and PDF exports.

- [#839](https://github.com/stella/folio/pull/839) [`13d93e6`](https://github.com/stella/folio/commit/13d93e6af7f8a4e7450858bc579033c22cad555e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render DOCX comment highlights through the shared display list and preserve comments as native PDF text annotations.

### Patch Changes

- [#840](https://github.com/stella/folio/pull/840) [`4918a4d`](https://github.com/stella/folio/commit/4918a4d7f3ab45e9aa704ef1e70b1a03988c1111) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render a bounded, non-editable preview for SmartArt diagrams while preserving the original OOXML for round-trips.

- [#838](https://github.com/stella/folio/pull/838) [`301b9e3`](https://github.com/stella/folio/commit/301b9e3442776c63c3a1a392ed6bd0bfbb18dd96) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve editable DrawingML WordArt metadata through parsing, editing, and DOCX serialization.
- Updated dependencies [[`301b9e3`](https://github.com/stella/folio/commit/301b9e3442776c63c3a1a392ed6bd0bfbb18dd96)]:
  - @stll/docx-core@0.20.1

## 0.38.0

### Minor Changes

- [#834](https://github.com/stella/folio/pull/834) [`89ee66d`](https://github.com/stella/folio/commit/89ee66d1c3aab8a83ea14f3c1f4c64f664d48f6b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Centralize inline-presentation semantics and use `FolioContentFormatRange` throughout comparison planning, results, and round-trip verification.

  Replace imports of the removed `CompareFormatRange` export with `FolioContentFormatRange` from `@stll/folio-core`.

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

### Patch Changes

- [#836](https://github.com/stella/folio/pull/836) [`e86528b`](https://github.com/stella/folio/commit/e86528b2c70aa2fbc30fdef4561dc1823bb52956) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose insertion controls, concrete numbering, and direct indentation consistently in agent operation schemas.

- [#836](https://github.com/stella/folio/pull/836) [`e86528b`](https://github.com/stella/folio/commit/e86528b2c70aa2fbc30fdef4561dc1823bb52956) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inline and display math inside tracked insertion, deletion, and move wrappers through parsing, editing, and serialization.

  Retain proofing exclusions in run formatting and fingerprint editable raw pictures so untouched DrawingML remains exact while model edits invalidate stale captures.

- Updated dependencies [[`e86528b`](https://github.com/stella/folio/commit/e86528b2c70aa2fbc30fdef4561dc1823bb52956), [`e86528b`](https://github.com/stella/folio/commit/e86528b2c70aa2fbc30fdef4561dc1823bb52956)]:
  - @stll/docx-core@0.20.0

## 0.37.5

### Patch Changes

- [#831](https://github.com/stella/folio/pull/831) [`d68669a`](https://github.com/stella/folio/commit/d68669a5fc9de91b5203d37023807f3a02b1c0a7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep block identity and fallback pairing inside compatible block families, containers, and table cells.

- [#830](https://github.com/stella/folio/pull/830) [`71b65f9`](https://github.com/stella/folio/commit/71b65f9a08e55eeaadfe2a4669ed4ed909ae94ab) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Centralize paragraph-property fallback emission in one exhaustive serializer.

- [#828](https://github.com/stella/folio/pull/828) [`8361198`](https://github.com/stella/folio/commit/836119895c23ced4c1924f3ee81775c093da14c5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Centralize DOCX text-formatting serialization so every serialization context shares one exhaustive implementation.

- [#832](https://github.com/stella/folio/pull/832) [`11e71d6`](https://github.com/stella/folio/commit/11e71d69aa3037835eceaf503f286720eca66155) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep captured paragraph properties aligned with modeled fallback emission when numbering provenance changes.

- [#833](https://github.com/stella/folio/pull/833) [`6a8a9b4`](https://github.com/stella/folio/commit/6a8a9b4fb303f2fcd34279614b1eda3a706394b5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep exact block anchors unique, nonblank, gap-local, and free of quadratic text matrices.

## 0.37.4

### Patch Changes

- [#825](https://github.com/stella/folio/pull/825) [`4e3a9dd`](https://github.com/stella/folio/commit/4e3a9dd520b4af75e97b575a9adf7b2a0ef1197d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve source-empty text box bodies without materializing editor placeholder paragraphs during save.

## 0.37.3

### Patch Changes

- [#815](https://github.com/stella/folio/pull/815) [`5276e94`](https://github.com/stella/folio/commit/5276e94e46232546fdbb1f92eb927fbb4541f50e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Adapt authored document fills in dark mode and fit oversized header and footer tables to their content frame.

- [#821](https://github.com/stella/folio/pull/821) [`c75a837`](https://github.com/stella/folio/commit/c75a837d2cac3557fbdd73017786c81dae6a9847) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Position and wrap column-relative text boxes within their active flow column.

- [#820](https://github.com/stella/folio/pull/820) [`99987b0`](https://github.com/stella/folio/commit/99987b0f083949296f3295b6c2ce10061a48d172) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve leading page breaks in styled paragraphs and table rows, classify row boundaries from their projected revision view, and coalesce repeated structural section carriers in linear time.

- [#819](https://github.com/stella/folio/pull/819) [`7385db9`](https://github.com/stella/folio/commit/7385db9807f293c57e1ca5eae072928bc45a93bf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep table-style fonts ahead of document-default fonts in table cells.

- [#819](https://github.com/stella/folio/pull/819) [`7385db9`](https://github.com/stella/folio/commit/7385db9807f293c57e1ca5eae072928bc45a93bf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and render DrawingML text-box rotation and flips across editor and PDF output.

- [#814](https://github.com/stella/folio/pull/814) [`bcc1ee4`](https://github.com/stella/folio/commit/bcc1ee46fcb36498c53bbf8dae8be37170f3dd2b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve opaque paragraph properties through collaborative DOCX edits with missing or duplicate paragraph IDs, including collapsed vertical-merge continuation cells.

- [#819](https://github.com/stella/folio/pull/819) [`7385db9`](https://github.com/stella/folio/commit/7385db9807f293c57e1ca5eae072928bc45a93bf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor page-margin header tabs without changing body tab stops.
- Updated dependencies [[`7385db9`](https://github.com/stella/folio/commit/7385db9807f293c57e1ca5eae072928bc45a93bf)]:
  - @stll/docx-core@0.19.4

## 0.37.2

### Patch Changes

- [#813](https://github.com/stella/folio/pull/813) [`fc9700b`](https://github.com/stella/folio/commit/fc9700bc835dd413e7359e043dfe8cc1b8baa21b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep edited table rows paired when another row is inserted or deleted.

- [#811](https://github.com/stella/folio/pull/811) [`8e77125`](https://github.com/stella/folio/commit/8e771259439fdcba1592bf3c4a6de24c20e805cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse immutable document story projections throughout DOCX comparison.

## 0.37.1

### Patch Changes

- [#809](https://github.com/stella/folio/pull/809) [`d094731`](https://github.com/stella/folio/commit/d094731b2c2cda733376cf0961ea10d13a3d4d8f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve full-document inline changes in one bounded headless transform.

- [#805](https://github.com/stella/folio/pull/805) [`aad2282`](https://github.com/stella/folio/commit/aad2282f11be2cae739737b3d1d97cae239a205e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph formatting when a comparison splits or merges paragraphs.

## 0.37.0

### Minor Changes

- [#795](https://github.com/stella/folio/pull/795) [`b3cc1d4`](https://github.com/stella/folio/commit/b3cc1d48038747c3001df03cf1bc874944f2b3cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Change format-painter capture and apply to use a typed effective-formatting value that preserves paragraph and character-style context, explicit inherited-property cancellations, and pending revisions.

- [#804](https://github.com/stella/folio/pull/804) [`a4c3df0`](https://github.com/stella/folio/commit/a4c3df0b0252babb700d75c09d8fbebb577a7d57) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a pure, resource-bounded `compareContent` API for ordered representation-neutral blocks, including text and formatting segments, moves, split and merge events, container-safe structural changes, and exact base and revised projections.

- [#799](https://github.com/stella/folio/pull/799) [`10ec72e`](https://github.com/stella/folio/commit/10ec72ea5d6742157f9683846ff95fd50ecfb0d0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit OOXML page-break run and field ownership across editor conversion, revision resolution, and DOCX round trips. Retain namespaces on nested run-property revisions and refuse page-break layouts that cannot be projected losslessly. AI edit snapshot anchors now require a structural-boundary fingerprint so edits cannot cross page-break topology changes.

- [#782](https://github.com/stella/folio/pull/782) [`09e9521`](https://github.com/stella/folio/commit/09e952108f2b420b542b280f9ba64bafa02abcbd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compare, inspect, insert, and track the complete direct paragraph-spacing cluster while preserving absent attributes, explicit zero and false values, and style inheritance.
  Line values and line rules retain independent direct-formatting provenance, so saving one no longer materializes an inherited counterpart.
  Paragraph-format operations now refuse a second unresolved serializable `w:pPrChange`; editor-only suggestion histories remain independently rejectable and save to at most one such child per paragraph.

### Patch Changes

- [#800](https://github.com/stella/folio/pull/800) [`2951504`](https://github.com/stella/folio/commit/295150452f9580d7dfb9a3133c7cdf7e92d2f1c6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep text geometry continuous across nested fragments and preserve style-linked list continuations.

- [#801](https://github.com/stella/folio/pull/801) [`975f804`](https://github.com/stella/folio/commit/975f8042b0128a8eee59fc623026ebf446dede53) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Anchor carets beside painted text when inline metadata has no visual span.

- [#803](https://github.com/stella/folio/pull/803) [`92d9533`](https://github.com/stella/folio/commit/92d9533f29dd7eb63e442bb66661da6349015f57) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep list deletion, paired inline metadata, markerless indentation, and rendered updates synchronized.

- [#779](https://github.com/stella/folio/pull/779) [`0aa34b2`](https://github.com/stella/folio/commit/0aa34b2e45efcb5a1eab9bdf3d2c356946057ea2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Limit resolved-story serialization errors to structural mismatch counts.

- [#787](https://github.com/stella/folio/pull/787) [`ea17fce`](https://github.com/stella/folio/commit/ea17fce2db901b24c0f2e7935ad8408478ff3d3c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve source part timestamps when normalizing paragraph identifiers so repeated normalization stays byte-deterministic.

- [#798](https://github.com/stella/folio/pull/798) [`7be42eb`](https://github.com/stella/folio/commit/7be42ebdd36e0cff7b913633ac52caa9c033351e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep table-grid widths aligned when accepting or directly applying column removals.

- [#795](https://github.com/stella/folio/pull/795) [`b3cc1d4`](https://github.com/stella/folio/commit/b3cc1d48038747c3001df03cf1bc874944f2b3cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Distinguish explicit off values from inherited boolean run formatting in comparison operations.

- [#783](https://github.com/stella/folio/pull/783) [`00dcefb`](https://github.com/stella/folio/commit/00dcefb66e888a16d2e1aa9b1288f07e326b07a7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Replace structurally incompatible table pairs as one deleted and one inserted table when their target template is lossless, preserving exact accept/reject round trips without sacrificing granular row edits or package-bound content.

- [#792](https://github.com/stella/folio/pull/792) [`f2f7574`](https://github.com/stella/folio/commit/f2f75740af14e438851e765278e10e2313694583) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve terminal table replacements across tracked-change acceptance, rejection, and reopen.

- [#804](https://github.com/stella/folio/pull/804) [`a4c3df0`](https://github.com/stella/folio/commit/a4c3df0b0252babb700d75c09d8fbebb577a7d57) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Retain and render word and formatting details on composite version changes, including edited moves and formatting-only scopes.

- [#795](https://github.com/stella/folio/pull/795) [`b3cc1d4`](https://github.com/stella/folio/commit/b3cc1d48038747c3001df03cf1bc874944f2b3cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve character-style toggle formatting consistently across body, header, footer, and footnote layout.

- [#797](https://github.com/stella/folio/pull/797) [`e289686`](https://github.com/stella/folio/commit/e2896863f28be23473628966a77123573bba2e0b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and recompute imported numbering across list structure edits.

- [#785](https://github.com/stella/folio/pull/785) [`2461ede`](https://github.com/stella/folio/commit/2461ededca2031784776e5ecaed00b7b0b6d0cac) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inherited paragraph spacing when inserted blocks are serialized, accepted, rejected, and reopened.

- [#786](https://github.com/stella/folio/pull/786) [`4c01126`](https://github.com/stella/folio/commit/4c0112677adf974d8ae7991733ef47abec1cf9d7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound the cost of resolving large tracked-change batches in synchronous headless review workflows.

- [#795](https://github.com/stella/folio/pull/795) [`b3cc1d4`](https://github.com/stella/folio/commit/b3cc1d48038747c3001df03cf1bc874944f2b3cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Distinguish authored run formatting from paragraph and character-style inheritance in reviewer snapshots.

- [#795](https://github.com/stella/folio/pull/795) [`b3cc1d4`](https://github.com/stella/folio/commit/b3cc1d48038747c3001df03cf1bc874944f2b3cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve direct run formatting, character-style inheritance, and formatted inline carriers (tabs, line breaks, symbols, and fields) across editor and DOCX round trips while keeping common run state compact and reviewable. Refuse a second independently owned run-property revision before mutation, and commit comment and revision ids only with the operation that writes them.

- [#787](https://github.com/stella/folio/pull/787) [`ea17fce`](https://github.com/stella/folio/commit/ea17fce2db901b24c0f2e7935ad8408478ff3d3c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve intended pagination around continuous sections, cached page markers, positioned objects, page furniture, and paragraph-mark formatting.

- [#794](https://github.com/stella/folio/pull/794) [`098cd24`](https://github.com/stella/folio/commit/098cd248c1f33f6849079360c1fb6f758927aa46) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve tracked section endpoints under paragraph-mark ownership.

- [#793](https://github.com/stella/folio/pull/793) [`8589725`](https://github.com/stella/folio/commit/8589725d71d29d77ca47f028188ac6787ed21055) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Paint floating page-furniture images at their authored positions.

- [#802](https://github.com/stella/folio/pull/802) [`68511cf`](https://github.com/stella/folio/commit/68511cf46320acf56722b0344853a8c996121302) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep list removal effective for paragraphs whose numbering comes from a style.

- [#788](https://github.com/stella/folio/pull/788) [`709ac7c`](https://github.com/stella/folio/commit/709ac7cad41fba216511fa1dbc678fda4e09fa43) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve and enumerate tracked insertions and deletions on inline atoms, including breaks and tabs, across document conversion, live suggestion editing, save, reload, and bulk resolution.

- [#791](https://github.com/stella/folio/pull/791) [`57a8099`](https://github.com/stella/folio/commit/57a8099d25bcd991f3be479eca0d28997654b3f6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve visible cell borders where a table row continues across pages.

- [#795](https://github.com/stella/folio/pull/795) [`b3cc1d4`](https://github.com/stella/folio/commit/b3cc1d48038747c3001df03cf1bc874944f2b3cf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restore rejected bulk run-property changes against their paragraph and character styles.
- Updated dependencies [[`e289686`](https://github.com/stella/folio/commit/e2896863f28be23473628966a77123573bba2e0b)]:
  - @stll/docx-core@0.19.3

## 0.36.0

### Minor Changes

- [#776](https://github.com/stella/folio/pull/776) [`6dcb318`](https://github.com/stella/folio/commit/6dcb3189516dd12728ec99c8b76f38b85f51cd14) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compare and track direct paragraph alignment changes while preserving the distinction between direct formatting and style inheritance.
  Paragraph replacement operations can now clear a direct paragraph style with `null`; paragraph insertion and property schemas expose their existing style and list clear values consistently.
  Unstamped multi-paragraph insert batches now reserve revision IDs for synthesized paragraph-property changes, so later batches cannot reuse an existing ID.
  Tracked paragraph insertion receipts include synthesized paragraph-property revisions, so targeted acceptance and rejection resolve the whole operation.
  Accepting or independently resolving suggested paragraphs at the end of a container keeps every final paragraph mark resolvable.

### Patch Changes

- [#769](https://github.com/stella/folio/pull/769) [`607c7b0`](https://github.com/stella/folio/commit/607c7b061ecec6d3f15dc6d82fc0ec1334e81887) Thanks [@jan-kubica](https://github.com/jan-kubica)! - An unchanged paragraph now writes its `w:pPr` properties back as they arrived, preserving validated unmodeled non-revision attributes and children and avoiding direct overrides synthesized from style-sourced numbering. The source is used only while a canonical formatting snapshot still matches and its structure passes validation; current section properties and tracked revisions are composed around it without accepting duplicates from the source. Unmodeled run-property revisions are excluded from captured full-repack replay until their accept/reject lifecycle is represented. Tracked-change UTC metadata now survives parser, editor, and serializer round trips under a canonical namespace binding.

- [#777](https://github.com/stella/folio/pull/777) [`4569d94`](https://github.com/stella/folio/commit/4569d94efb1ca031e6c2182fcc8af4d31338f388) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep unique terms aligned without inventing opposite-direction edits. Long
  paragraphs factor common text and unique anchors before bounded LCS work, and
  one compact work allowance now covers each document comparison or apply batch.
  Atomic preflight preserves that allowance with a coarse zero-DP check.
  Normalized comparison-key storage is capped before allocation. The generated
  declaration budget rises by 15 lines for the internal shared-session helpers.

## 0.35.1

### Patch Changes

- [#773](https://github.com/stella/folio/pull/773) [`edb6ff1`](https://github.com/stella/folio/commit/edb6ff1bd5bbe2feb9446202aec1ae6e525d2e9c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve run property revision boundaries when parsing DOCX content.

- [#775](https://github.com/stella/folio/pull/775) [`3c90955`](https://github.com/stella/folio/commit/3c90955f7e61215b7479747e25067bf761f2f7c6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Widen the `@stll/template-conditions` dependency to `>=0.4.0 <1.0.0`. folio-core consumes only the scanner surface, so a host monorepo that already provides template-conditions as a workspace package satisfies the range across 0.x minors instead of installing a second registry copy beside its own.

## 0.35.0

### Minor Changes

- [#771](https://github.com/stella/folio/pull/771) [`9d63dfe`](https://github.com/stella/folio/commit/9d63dfec2846df376084e6dd756b0e0d7982d909) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Template directives follow the docxtpl dialect of Jinja.

  `@stll/folio-core` now scans markers with `@stll/template-conditions` 0.4:
  `{{ path | filter(...) }}`, `{% if %}` / `{% elif %}` / `{% else %}` /
  `{% endif %}`, `{% for alias in path %}` / `{% endfor %}`,
  `{{ clause("Name") }}`, `{{ num("key") }}`, `{{ ref("key") }}`, and the
  `{{ loop.* }}` counters.

  `DirectiveKind` renames accordingly (`each` → `for`, `endeach` → `endfor`,
  `elseif` → `elif`, `index`/`count` → the single `loop` kind). A `for`
  `DirectiveRange` carries the iterated array path in `expr` and the loop alias in
  the new optional `alias` field.

  The React and Vue overlays rename the kind-derived class suffixes to match:
  `--each` → `--for`, `--endeach` → `--endfor`, `--elseif` → `--elif` on
  `.folio-template-directive`, `--each` → `--for` on `.folio-template-band-rail`,
  plus a new `.folio-template-directive--loop`. Closer hover hints read
  `endif · <opener expression>` / `endfor · <opener expression>`.

### Patch Changes

- [#768](https://github.com/stella/folio/pull/768) [`a9edc21`](https://github.com/stella/folio/commit/a9edc218902b265527c54c735f2ed4bca16566ce) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Report numbering changes only for levels referenced by a paragraph in either document.

## 0.34.0

### Minor Changes

- [#756](https://github.com/stella/folio/pull/756) [`0a90bec`](https://github.com/stella/folio/commit/0a90becf8317745f447df426b5c3f6be626c1300) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compare and track direct font family, half-point size, and RGB color changes without rewriting unchanged text.

### Patch Changes

- [#758](https://github.com/stella/folio/pull/758) [`f16a9ba`](https://github.com/stella/folio/commit/f16a9baa5df35e57bfb22d4157512aa18ff54fbf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render legacy form checkboxes without cached field results.

- [#759](https://github.com/stella/folio/pull/759) [`62645bc`](https://github.com/stella/folio/commit/62645bcf83ed313a58b66b6b6eff274e5f75c394) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep bounded final words on modern indented justified lines.

- [#748](https://github.com/stella/folio/pull/748) [`a87b719`](https://github.com/stella/folio/commit/a87b719519733f22057f9aeeac05b6d519544eca) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use David's font metrics when calculating automatic line spacing.

- [#750](https://github.com/stella/folio/pull/750) [`7a5737a`](https://github.com/stella/folio/commit/7a5737a9616c6fbb0e61ca02f3e5e2a0c63832aa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve FrankRuehl and Miriam with compatible Hebrew fallbacks and verified line metrics.

- [#753](https://github.com/stella/folio/pull/753) [`52c0a4b`](https://github.com/stella/folio/commit/52c0a4bd7f4733afa40983b0508560a0a0c1b6d5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Scope default line-edge punctuation restrictions to East Asian text.

- [#749](https://github.com/stella/folio/pull/749) [`6510158`](https://github.com/stella/folio/commit/6510158681e67385b988045485a2c03ddae6f068) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Apply odd, even, and continuous section starts at their preceding section boundaries.

- [#743](https://github.com/stella/folio/pull/743) [`b7fa6d1`](https://github.com/stella/folio/commit/b7fa6d18f10c7058e4e3445549889ad4856b38a0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve watermark hosting paragraphs when they contribute header clearance.

- [#765](https://github.com/stella/folio/pull/765) [`129d020`](https://github.com/stella/folio/commit/129d020b67f69992fdede898f2c637ae6a741d81) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resize shape-fitted text boxes to their measured content height.

- [#761](https://github.com/stella/folio/pull/761) [`34291e1`](https://github.com/stella/folio/commit/34291e1f9439f5a990ca6cb07804161fdea8b4a2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render endnote reference markers with their configured number format and the standard lower-Roman default.

- [#760](https://github.com/stella/folio/pull/760) [`b2426ae`](https://github.com/stella/folio/commit/b2426ae3c218730e4554d1913d317d3ee371bf40) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render package-owned picture watermarks at their authored shape size.

- [#744](https://github.com/stella/folio/pull/744) [`7bed19e`](https://github.com/stella/folio/commit/7bed19ea6a95f769772ed3ed573d17b2727da28c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor cached page boundaries on exact-height multi-cell table rows.

- [#763](https://github.com/stella/folio/pull/763) [`107682a`](https://github.com/stella/folio/commit/107682a26b6540e617af2ea63c849f9c3d80d804) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve the height of authored terminal footnote paragraphs after tables.

- [#747](https://github.com/stella/folio/pull/747) [`6451982`](https://github.com/stella/folio/commit/645198226e696e95e953515fb94f2d624101f095) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Exclude fully hidden paragraphs from visual layout while preserving their editable source runs.

- [#745](https://github.com/stella/folio/pull/745) [`5748012`](https://github.com/stella/folio/commit/574801248f096838c7b3fbbd84b79d456b105970) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve header clearance for unformatted watermark host paragraphs.

- [#746](https://github.com/stella/folio/pull/746) [`0b285c3`](https://github.com/stella/folio/commit/0b285c3f74d56f7cb7fbb648bc44d2ac23f0485e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render percentage shading patterns as blended foreground and background colors.

- [#752](https://github.com/stella/folio/pull/752) [`24da558`](https://github.com/stella/folio/commit/24da558aa554cbcb43ddad3f6145febcc011c1a7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Write a section break where the schema places it in a paragraph's properties: after the mark's run properties and before a recorded property change. A paragraph that ended a section and also recorded a property change had its break written after the change, and a consumer refused the part.

- [#754](https://github.com/stella/folio/pull/754) [`954f443`](https://github.com/stella/folio/commit/954f44375303d41b20b132ecf745a689256738f3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor paragraph and conditional-region precedence in table style cascades.

- [#755](https://github.com/stella/folio/pull/755) [`17d193b`](https://github.com/stella/folio/commit/17d193b95b98da8b1fc16dc415144162e0a07797) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve bulk tracked changes with fewer document steps.

- [#751](https://github.com/stella/folio/pull/751) [`27e2717`](https://github.com/stella/folio/commit/27e2717fea8dbd367198fdc520ed5ab3364d9828) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Apply section line grids to table cells when the document compatibility setting requests it.

- [#766](https://github.com/stella/folio/pull/766) [`e44817d`](https://github.com/stella/folio/commit/e44817dd83a4ed34cffb2e8846ff7cd50789e114) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit end-tab positions inside the content frame when a paragraph has a right indent.

- [#764](https://github.com/stella/folio/pull/764) [`082433e`](https://github.com/stella/folio/commit/082433e25e16930f05956f39a5c284b7dfa9249d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Position floating text boxes from their exact host paragraph.

- [#770](https://github.com/stella/folio/pull/770) [`5b4795b`](https://github.com/stella/folio/commit/5b4795b4a121d689ca28c4daf8b243ba5fbae426) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve watermark host clearance across incremental and text box layout paths.

- [#762](https://github.com/stella/folio/pull/762) [`8b9fc7d`](https://github.com/stella/folio/commit/8b9fc7de68668b23d9dca6634e0bde5335b7fc6b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve layout space for relationship-backed images whose format cannot be painted.
- Updated dependencies [[`27e2717`](https://github.com/stella/folio/commit/27e2717fea8dbd367198fdc520ed5ab3364d9828)]:
  - @stll/docx-core@0.19.2

## 0.33.2

### Patch Changes

- [#740](https://github.com/stella/folio/pull/740) [`a7539ff`](https://github.com/stella/folio/commit/a7539ffcb62b7a2248fce9cf0cffe356bc1001ac) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Declare every namespace prefix a rebuilt part uses. `word/document.xml`, headers, footers, note parts, `comments.xml`, `commentsExtended.xml` and the numbering, styles, settings, font-table and theme parts now derive their `xmlns:*` from the markup they emit instead of a hand-maintained list: each prefix is resolved through one namespace table, falling back to the bindings the source part's root declared, and a prefix that resolves to nothing fails the save with an `UnboundNamespacePrefixError` rather than being written unbound. `mc:Ignorable` comes from the same table, so it names only prefixes the part declares and every declared extension prefix it should. A document carrying content the parser preserves verbatim — a text box under `wne:txbxContent`, a `w16du:dateUtc` revision stamp — no longer saves to a part a consumer refuses to open.

- [#734](https://github.com/stella/folio/pull/734) [`2e4677b`](https://github.com/stella/folio/commit/2e4677b3df668ddff7f1b07a43442b6ac6128fc5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compare: a container's final paragraph mark now carries no revision in either direction. Paragraphs added where a removed trailing run was land in the paragraph the container ends with; paragraphs appended past it rotate the added break one paragraph back; and a carrier no merge chain can reach is reserved for the target's last paragraph. Paragraph properties no longer overwrite a paragraph mark another operation in the same batch wrote, a paragraph followed only by a text box is read as ending its container, `w:trPr/w:del` rows keep their cells' marks, every story is compared as accepted, and `w:pPr` children are written in schema order.

- [#733](https://github.com/stella/folio/pull/733) [`fc0460b`](https://github.com/stella/folio/commit/fc0460bf7cfd17cdb54daea0d200e1c530589a4b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Format the final-paragraph-mark guard's container path only where it finds one, so the walk costs no more than the projection beside it.

- [#735](https://github.com/stella/folio/pull/735) [`cb2b28d`](https://github.com/stella/folio/commit/cb2b28d6844ad6fd4b1b37bebf0a25751fd32b11) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry a table's own properties through `compareDocx`: an inserted table brings the target's `w:tblPr`, `w:tblGrid` widths, `w:trPr` and `w:tcPr` (spans, merges, shading, borders, margins, alignment) and its nested tables; a removed table keeps them under the deletion marks, with the runs inside a deleted row marked too; a paired table, row or cell whose properties changed records `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange`. The round-trip self-check now compares the table model as well as the blocks, with a `table-geometry` cause. `w:tblPr` children are written in the order `CT_TblPrBase` declares, so a table carrying both `w:tblLayout` and `w:tblCellMar` validates.

- [#737](https://github.com/stella/folio/pull/737) [`05e7edb`](https://github.com/stella/folio/commit/05e7edbd1876910fc73368530bdab8859696be32) Thanks [@jan-kubica](https://github.com/jan-kubica)! - State an application version of the form the schema fixes on every save. The extended-properties `AppVersion` in `docProps/app.xml` is `XX.YYYY`, and a value with two dots is refused by consumers outright; folio copied the part through verbatim when saving a document it had not created, so a package could carry such a value in and back out. Both save exits — the full repack and the selective save — now rewrite a value the form rejects to one derived from it alone (`1.0.0` becomes `1.0000`), leave the rest of the part byte-identical, and add extended properties to no package that lacked them.

- [#739](https://github.com/stella/folio/pull/739) [`af99832`](https://github.com/stella/folio/commit/af998329b6676a2eac4b421c0a252abb271e3f01) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Export `currentFolioBlockId`: the id a block recorded under an earlier version answers to now. A paragraph id above the 31-bit bound is brought into range on parse, so a block id recorded from such a paragraph before that stopped resolving; the mapping is a pure function of the id, so a host resolves a recorded id without the document.

- [#742](https://github.com/stella/folio/pull/742) [`173c774`](https://github.com/stella/folio/commit/173c7744bb42ffaf975f381baaf82aeae5be7a9c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A rebuilt package is Transitional throughout. folio reads both conformance classes and writes only Transitional, so content the parser preserves verbatim — raw property XML, drawing and shape bodies, text-box markup, unmodeled extensions — used to reach a rebuilt part in the source's own spelling: a Strict fragment kept its Strict namespace declarations under a Transitional root, and its lengths and percentages stayed in the form that carries its unit (`155.85pt`, `20%`) where the Transitional attribute is a number, which a validating consumer rejects on `w:tcW`, `w:tblW` and the size-relative drawing extensions. Every capture now goes through one owner that converts the fragment where its namespace scope still says which class produced it, and a lint rule keeps the raw serializer out of the parsers so a replay path added later cannot skip it. The attributes that need converting are derived from the repository's Transitional schema graph rather than listed by hand, with the unit each type counts in stated once and checked in CI; a Strict namespace with no Transitional counterpart fails the save instead of being written. The typed projection reads a length through the same table, so a Strict table no longer lays out at a twentieth of its width, and a shape with no text body writes the `wps:bodyPr` its content model requires.

- [#738](https://github.com/stella/folio/pull/738) [`da95224`](https://github.com/stella/folio/commit/da9522486bdfd5e50c5fa935f10ebb3173b06cd6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A save that changed nothing writes a table's `w:tblPr`, `w:tblGrid`, `w:trPr` and `w:tcPr` back as they arrived, rather than rebuilding them from the typed model and dropping the conditional-format flags, the `w:tblGridChange`, and whatever else the model does not cover. The capture is re-parsed and checked against the model before it is used, so a `Document` edited in place is still honoured. `w:tcPr` and `w:tblPr` also stop gaining an inherited value — a border a table style supplied, a margin the table declared — as the cell's or table's own override, and an absent `w:hideMark` stops being written back as an explicit `w:val="off"`.
- Updated dependencies [[`da95224`](https://github.com/stella/folio/commit/da9522486bdfd5e50c5fa935f10ebb3173b06cd6)]:
  - @stll/docx-core@0.19.1

## 0.33.1

### Patch Changes

- [#730](https://github.com/stella/folio/pull/730) [`3cd70f9`](https://github.com/stella/folio/commit/3cd70f9124c5cf84597e7a400750773b2da93225) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep a container's final paragraph mark when a comparison removes the paragraphs it ends with: the merge chain now runs from the last surviving paragraph forward, and the carrier's properties travel as `w:pPrChange`.

- [#731](https://github.com/stella/folio/pull/731) [`4daa0a5`](https://github.com/stella/folio/commit/4daa0a574a99aad6edeeba471ad86f7b919b69c8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Emit schema-valid OOXML for compared packages, and date them from the comparison.

  Four shapes a compared package could carry were not shapes the schema allows,
  and the round-trip self-check could not see any of them: it reads which words a
  view resolves to, and all four parse back to the right words.

  - **A revision `w:id` is unique across the package.** One logical change
    serializes as several physical wrappers — a word-level redline cut around the
    words that survived — and the pass that gives each its own id ran only on the
    full repack. The selective save, which rewrites the changed paragraphs and
    re-emits every other part verbatim, exited past it, so a paragraph-local edit
    shipped several revisions under one id. Both exits now run the pass, and every
    id it lets stand or mints goes through one choke point that refuses a
    duplicate.
  - **`w:tbl` is `w:tblPr`, `w:tblGrid`, then rows.** Both are required and both
    precede every row; they were emitted only when the model had something to put
    in them, so a table the comparison creates — no authored properties, no
    measured column widths — opened with its first `w:tr`. They are now always
    written, with a grid column per column the widest row spans.
  - **`w:hyperlink` wraps a revision, not the other way round.** Run-level
    `w:ins`/`w:del` take run-level content, which a hyperlink is not. Deleting or
    inserting linked text now emits
    `<w:hyperlink><w:del><w:r><w:delText>…`, splitting the revision at each link
    boundary, and reading such a package back restores the same model.
  - **A paragraph id is 31-bit.** `w14:paraId`, `w14:textId` and the comment-part
    ids that reference a paragraph are `ST_LongHexNumber` with a maximum below
    `0x80000000`. Producers exist that ignore the bound, and folio preserves the
    ids a document arrives with, so an out-of-range id travelled straight through
    a save. One mapping, a pure function of the id, now brings such a value into
    range — applied when a paragraph is parsed and again across every part of the
    package on the way out, so a paragraph and every reference to it move
    together and a document's identity does not shift between reading and
    writing.

  `compareDocx` also restamps `dcterms:modified` in `docProps/core.xml` from its
  `timestamp` option. The save wrote the wall clock there, so two runs over
  identical inputs with identical options differed in that part alone — the last
  clock in a call whose whole contract is that it has none.

## 0.33.0

### Minor Changes

- [#713](https://github.com/stella/folio/pull/713) [`3c46347`](https://github.com/stella/folio/commit/3c463478e19fd21f5f9a526fcc8288a66b59d97f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Build a display page when it is painted, and let a run say which document its
  positions belong to.

  The editor's display-list renderer built the whole document's list on every
  layout run, then painted the three pages on screen from it. Measured on a
  ninety-one-page document, that cost 49.6 ms per painted page against the
  painter's 0.567 ms: work proportional to the document where the page container
  does work proportional to the screen. `createDisplayListBuilder` computes what
  belongs to the document once and each page when it is asked for, which brings
  the same measurement to 0.43 ms per painted page.

  `DisplayGlyphRun.pmRange` now carries the story it addresses, so a header, a
  footer and a note run can carry one at all. A page is not one document: the same
  position means a different character in the body, in each header and footer
  part, and in each note, so a range that did not name its story could only be
  used for the body. The producer used to drop the others and the painter used to
  strip them; both now travel, named, and the DOM backend writes them onto the
  element that paints the run.

- [#715](https://github.com/stella/folio/pull/715) [`55b5f43`](https://github.com/stella/folio/commit/55b5f43b917191303287af9a05e961486026d18f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A page says where a click can land, not just what it paints.

  `DisplayPage` gains `regions`: a tree of boxes over the same primitives — the
  content area, header and footer slots, notes, paragraphs, lines, empty lines,
  tables, rows and cells — each carrying the model range and the story it
  resolves to, plus the block id, comment threads and row and column indices a
  surface reads. The producer opens a region around the painting that fills it, so
  the structure comes from the walk that lays the page out rather than from a
  second pass over it, and a region indexes into the paint list rather than
  copying it.

  The DOM backend paints each region as the element the interaction layer has
  always looked for, with the runs nested inside, so clicks, drags and selections
  resolve against what the producer laid out. A glyph run also states when it is a
  line-edge space run the line was fitted without, and what one of those spaces
  would have advanced, which is what lets a caret step through them.

  `interactionContract.test.ts` reads the interaction layer's own source, extracts
  every class and data attribute it looks for, and fails unless each one has a
  source in the IR. A reader that learns a new selector without the producer
  gaining something to emit it from fails that test.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `insertTable` and `deleteTable` add and remove a whole table: every row marked
  `w:trIns` or `w:trDel` in tracked mode, which is how Word says it. `compareDocx`
  emits them as `table-insert` / `table-delete`, so a pair whose table count
  differs is compared instead of refused.

  Its segment pairing no longer goes by index. A table pairs with the table
  opposite it unless the next table on one side matches it better, so removing
  the first table no longer shifts every later one and rewrites each table's
  contents into the next. A segment is the outermost table plus everything nested
  in it; `FolioAIBlock.table` gains `outerTableIndex` to carry that distinction.

- [#706](https://github.com/stella/folio/pull/706) [`d7962d1`](https://github.com/stella/folio/commit/d7962d1c41f990260e3f2b81aef8bfd89dadfa62) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `compareDocx` now reports what its self-check found instead of only acting on
  it, and can be asked for its best attempt when it cannot prove one.

  Every successful result carries `verification`: `{ status: "verified" }`, or
  `{ status: "unverified", failures }` where each failure names the invariant that
  did not hold (`accept-reproduces-target` or `reject-reproduces-base`), the
  projection field that diverged (`container`, `block-count`, `style`,
  `list-level`, `invisible-structure`, `whitespace`, `text`), the story it
  happened in, and a structural detail carrying no phrase of either document.

  The default is unchanged: an unproven redline is refused, because a reader
  cannot tell one that lost something from one that did not.
  `CompareDocxRoundTripError` now names the invariant and the cause and carries
  the whole failure list, in place of the two block-text arrays it used to hold.

  `onUnverified: "emit"` is the opt-in for the other trade — the redline it could
  build, plus the typed list of what it could not represent. A parse, apply or
  serialize failure is still an error under either setting: there is no redline
  to emit.

  Both directions of the round trip are now checked. The self-check previously
  proved only that accepting reproduces the target; it also proves that rejecting
  reproduces the base it was written against.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `diffWordSegments` stops shredding a rewritten paragraph. An LCS maximises
  matched characters, so a rewritten sentence used to come back as a dozen
  struck-through fragments interleaved with a dozen inserted ones. Three rules
  pull it back: a match made only of separators is not a match, a one-token match
  with changes on both sides of it is dropped into them, and a paragraph whose
  surviving matches are too short for its length is replaced whole. On a
  320-paragraph rewrite the package carries 68% fewer separately marked runs; a
  light edit is marked word by word exactly as before.

  The diff now takes options: `granularity` (`"word"` default, or `"character"`
  to mark the changed letters inside a token) and `normalization` (`case`,
  `whitespace`). `granularity` is threaded through `applyFolioDocumentOperations`
  and `compareDocx`; normalization is not, because a comparison that leaves a
  difference unmarked does not accept back to the target.

- [#705](https://github.com/stella/folio/pull/705) [`627e8bc`](https://github.com/stella/folio/commit/627e8bc24cb79f1ec8643150cf90d3fbe2c2cf0d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Paint the editor's pages from the display list, behind a renderer option.

  `buildDisplayList` now takes the page furniture it previously could only
  report: page borders, watermarks, footnote bodies, header and footer stories,
  and the package's own embedded font faces. A construct that is supplied is
  painted; one the document has but the caller withheld is still reported; one
  the document does not have is neither. `layoutDocxHeadless` produces all of it,
  so an export paints the pages an editor paints rather than bare bodies.

  The DOM backend places every code point at the advance the layout engine
  measured instead of letting inline layout advance it, so the two backends agree
  on glyph positions to within the browser's 1/64 px layout quantum. Cursively
  joined clusters stay in one box, because only shaping can choose a positional
  form; the advances inside such a cluster are the shaper's.

  `pageRenderer` on the React and Vue editors selects which renderer paints the
  pages, defaulting to the existing painter. Only painting is swapped: page
  shells, virtualization, the fingerprint comparison that skips an unchanged page
  and the painted event stay shared, so incremental repaint behaves the same
  under either renderer.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `compareDocx` reports a changed numbering definition as a `numbering` change:
  a list whose format, level template or start differs moves every label in the
  list and no block's text, so a text comparison saw two identical documents.
  Reported and not represented — OOXML has no tracked-change grammar for
  `numbering.xml`, and Word does not track it either.

  `FolioDocxReviewer.readNumberingDefinitions()` returns the package's numbering
  flattened to one entry per instance and level, with overrides resolved.

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

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `setBlockParagraphProperties` changes a block's list level or paragraph style
  without touching its words, recorded as a `w:pPrChange` carrying the complete
  previous property set so a reject restores it the way Word does. `FolioAIBlock`
  gains `listLevel`, and the two insert operations take `listLevel` and a
  nullable `styleId` so an inserted paragraph no longer takes its level and style
  from whichever block happens to follow it.

  `compareDocx` reports the edit as `paragraph-format`. A demoted list item used
  to reach the comparison as no change at all, so the redline said two different
  documents agreed. Its round-trip self-check now covers each block's style and
  list level alongside its text.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `compareDocx` compares every story present on both sides — main, headers,
  footers, footnotes and endnotes — instead of the body alone. A pair differing
  only in a footnote used to be reported as agreeing.

  `FolioDocumentOperationResult` gains `nextRevisionId`: the first revision id a
  following batch may allocate against the same document. The revision-id space
  is the package rather than the part, so a caller writing one batch per
  story has to seed each from the previous batch's value; the batch is the only
  thing that knows how many ids it took.

  `FolioDocxReviewer.acceptAll` and `rejectAll` now sweep every story rather than
  the body, so a revision in a header or a note no longer survives an accept-all.

  `COMPARE_UNSUPPORTED_REASONS` drops `secondary-story` and gains
  `story-not-editable`; a story present on one side only is still reported as
  `story-missing-in-base` / `story-missing-in-target`.

- [#694](https://github.com/stella/folio/pull/694) [`6b0e4a6`](https://github.com/stella/folio/commit/6b0e4a6cc387965cf7cdb708a176650289574edf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `compareDocx(base, target, { author, timestamp })`: a deterministic
  two-document compare returning the base package with tracked changes that
  accept back to the target and reject back to the base, plus a JSON change list
  (`insert`, `delete`, `replace`, `move`, `format`, `table-row-insert`,
  `table-row-delete`). Header, footer, footnote, and endnote stories are reported
  as unsupported rather than silently skipped.

  The call reads no clock and no random source, so the same inputs give
  byte-identical output. Supporting that, `FolioRevisionStamp` lets any apply
  batch pin its revision date and id seed, and `FolioAIBlock.table` records the
  block's enclosing table cell.

- [#710](https://github.com/stella/folio/pull/710) [`4974a68`](https://github.com/stella/folio/commit/4974a68a62d93d2853f4419f6c16ef74390c57b3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Carry what the measurer applied on the glyph run, so a backend can reproduce
  the width the line was fitted at.

  `advancesPx` folded letter spacing, a horizontal scale, justification, kerning
  and small capitals into one number per code point. A backend that hands the run
  to a shaper cannot recover any of them from the text: the shaper advances glyphs
  by what the font says, and none of those five is in the font. The DOM backend
  therefore painted runs at the glyphs' own width rather than the laid-out one,
  by as much as 78 px on a justified line.

  `DisplayGlyphRun` now names them: `adjustments` carries the letter spacing,
  horizontal scale and per-space justification delta, and `kerning` and
  `smallCaps` state what the advances were measured with. `DisplayFontFace` gains
  `fallbacks`, the families between the first and the generic, because a face is a
  stack and a backend handed only its first entry paints a different face from the
  one measured wherever that entry is missing.

  A run's advances now also sum to the width the line was broken on. They were the
  sum of per-character measurements, which differs from the string's own width
  wherever a pair kerns or ligates; the difference is spread across the run rather
  than left as an extent no backend paints.

  The editor's run-drift check gates the painted extent as well as the origin:
  every run in the two fixtures now lands within one browser layout quantum.

- [#709](https://github.com/stella/folio/pull/709) [`916b84d`](https://github.com/stella/folio/commit/916b84def4f48f476e74b77641b9d44bf92b8799) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Shape text in the scripts that need it, from one implementation both the
  measurer and the PDF backend use.

  Arabic, Hebrew, the Indic scripts and their neighbours do not select one glyph
  per code point: a letter takes its form from the letters beside it, lam followed
  by alef ligates, a Devanagari cluster reorders into a conjunct, a Hebrew point
  hangs off the letter it belongs to. A new `stella-text-shaper` crate answers
  that question over rustybuzz, and `packages/core/src/shaping` is the only way to
  it, so a measurement in CSS pixels and a PDF text matrix scale the same glyph
  ids and the same advances.

  The headless measure provider now measures such a run by its clusters rather
  than a code point at a time, and the PDF backend paints the glyphs shaping
  chose, subsets them, and maps each back to the characters that formed it so a
  ligature or a conjunct still extracts as text. `writePdf` is asynchronous as a
  result, and no longer reports `unshaped` runs or refuses them under
  `strictShapedScripts`: the runs it used to name are painted correctly.

  The shaper is a separate WebAssembly artifact with its own size budget, fetched
  the first time a document actually contains a run that needs it. A document in
  Latin, Cyrillic or Greek never loads it.

- [#729](https://github.com/stella/folio/pull/729) [`9b3defa`](https://github.com/stella/folio/commit/9b3defaa25d805b04143fa174a7d142860db7c21) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Let a host say which key presses the editor's page-level shortcuts answer.

  `DocxEditor` takes a `keyboardShortcuts` prop: `"document"` (the unchanged
  default) answers every press on the page, `"editor"` answers only a press
  landing inside the editor, and `"none"` binds no page-level listener at all.
  A host that docks the editor beside its own panes keeps its own bindings and
  opens the dialog through the new `DocxEditorRef.openFind` / `openReplace`,
  which seed the search box from the current selection exactly as Cmd/Ctrl+F
  does. The scope predicate is `isKeydownInShortcutScope` in
  `@stll/folio-core/managers/editorShortcuts`; `useWheelZoom` takes the same
  scope in place of its `enableKeyboardShortcuts` flag.
  The compat `DocxEditor` forwards its legacy `disableFindReplaceShortcuts` flag
  as `keyboardShortcuts: "none"` instead of dropping it.

- [#703](https://github.com/stella/folio/pull/703) [`2473407`](https://github.com/stella/folio/commit/24734075bca57bc54a766cd090a4a1a0a633b54a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a painter-neutral display list and a native PDF backend that consumes it.

  `buildDisplayList` turns a laid-out document into an ordered list of paint
  primitives per page: glyph runs carrying the advances layout was decided on,
  filled and stroked rects, lines, images, clip, rotate and opacity groups, link
  annotations and a heading outline. Two backends consume it and nothing else:
  `renderDisplayListToDom` paints it into DOM elements, and `writePdf` writes a
  PDF with subset TrueType faces, PNG and JPEG images, vector borders and
  shading, hyperlinks and an outline. A backend that reads layout data the
  display list does not carry now fails a dependency-cruiser rule rather than a
  review.

  `exportDocxToPdf` composes the whole chain without a browser:
  `layoutDocxHeadless` paginates a package through the measurement seam, and
  `installHeadlessMeasureProvider` supplies that seam from parsed font binaries
  instead of a canvas. Output is deterministic: `timestamp` is required rather
  than defaulted, so two exports of one document are byte-identical.

  A font source supplies every binary that carries part of a face, not one, and
  both measurement and embedding resolve each code point to the binary that
  covers it. Families are routinely shipped split by script, so a Czech, Slovak
  or Polish document needs two subsets of one family in the same paragraph;
  resolving per face rather than per code point would paint an empty box for
  every character outside whichever subset was chosen. A code point no supplied
  binary can encode is reported in `unencodable` rather than painted silently,
  and `strictGlyphCoverage` turns it into a failure for a caller who would
  rather not ship the page at all.

  The existing layout painter is unchanged and still paints the editor. The
  display-list DOM backend is additive in this release.

- [#723](https://github.com/stella/folio/pull/723) [`006ba65`](https://github.com/stella/folio/commit/006ba65bf738695625cc3a169ba048f12777a75f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Compare strikethrough as a tracked inline-formatting change.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Two operations for the edit that moves a paragraph mark and no words:
  `splitBlock` writes an inserted mark on the paragraph the break now ends, and
  `mergeBlockWithNext` a deleted one. Both carry the `separator` the break stands
  in for, deletion-marked on a split and insertion-marked on a merge, so either
  direction of accept/reject reproduces the right spacing. A deleted mark is
  refused where there is no sibling to join with — the last paragraph of a table
  cell, or of a story.

  `compareDocx` emits them, so a split is reported as `split` and a merge as
  `merge` rather than as a rewrite of the half that stayed put plus an insertion
  or deletion of the other. On a 320-paragraph document of splits and merges the
  change list drops from 145 entries to 87 and the redline's text-carrying runs
  from 156 to 87.

- [#720](https://github.com/stella/folio/pull/720) [`b489528`](https://github.com/stella/folio/commit/b489528bdce66a9c6215eee90e19b2afb80cb283) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align uniquely matching table columns and report tracked column insertions and deletions.

- [#721](https://github.com/stella/folio/pull/721) [`681923a`](https://github.com/stella/folio/commit/681923ab277b78acc69f3eeaea0262ec07fcf168) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve bookmark boundaries inside tracked inline changes so accepting and rejecting revisions keeps bookmark ownership intact.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A relocation is written into the document as a linked pair.
  `deleteBlock`, `insertAfterBlock` and `insertBeforeBlock` take a `moveId`; when
  one id names exactly one deletion and one insertion the applier writes
  `w:moveFrom` and `w:moveTo` instead of an unrelated deletion and insertion.
  An id that does not is reported as an `unpairedMove` normalization and both
  halves apply plainly.

  `compareDocx` emits the pair, so a reordered document now says so to every
  OOXML consumer rather than only in its JSON change list. A relocated paragraph
  is recognized when it keeps at least 80% of its word tokens, so a clause edited
  on the way to its new home is still a move.

  `FolioAIEditNormalization` is a discriminated union on `code` rather than one
  shape with a `splitMultilineText`-specific field.
  `FolioDocumentOperationResult.nextRevisionId` is optional: a host bridge that
  delegates to an editor it does not control omits it rather than guessing.

### Patch Changes

- [#712](https://github.com/stella/folio/pull/712) [`f3b1f14`](https://github.com/stella/folio/commit/f3b1f1456af587f12b9a9a23cb27932136fda4bb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Stop reporting a watermark-only header as a story that never arrived.

  Word puts a watermark in a header part that holds nothing else, so that
  header converts to no paintable content and the display list reported it as
  "the header part this page selects was not among the supplied stories" on
  every page of every watermarked document. The watermark painted from that same
  part proves it reached the producer and was read, so it is no longer reported
  as a gap; a header that names content the builder really did not get still is.

- [#717](https://github.com/stella/folio/pull/717) [`79ac7cb`](https://github.com/stella/folio/commit/79ac7cbcd1f60f2bc4bd3153c98056a62a9c6685) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve editing hit regions inside clipped table rows.

- [#722](https://github.com/stella/folio/pull/722) [`16f8546`](https://github.com/stella/folio/commit/16f85467f9f4111d2715adaac7aea46d8a07b87e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph style and list level in generated move scenarios.

- [#728](https://github.com/stella/folio/pull/728) [`4400948`](https://github.com/stella/folio/commit/4400948699fb154a55fe2b8f3fdd595fb9f21186) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the required final body paragraph outside tracked paragraph-mark deletions.

- [#697](https://github.com/stella/folio/pull/697) [`2f6e5eb`](https://github.com/stella/folio/commit/2f6e5ebbc8abf3d2c541dcfab3e7643f2140a9cb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A save now carries every part of the source package byte for byte — embeddings,
  media, custom XML, fonts, macro projects, ActiveX controls, parts folio does not
  model — and a macro-enabled document keeps its main-part content type. The only
  entry a save refuses is one whose path would escape the package, and its
  relationships and content-type entries leave with it, so a saved package never
  references a part it no longer holds.

- [#716](https://github.com/stella/folio/pull/716) [`3408880`](https://github.com/stella/folio/commit/3408880b0b7945b76addfd744fe31395299f50e7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the last word of a justified line that fills its measure exactly in documents that predate the current justification rules.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `compareDocx` compares the accepted view of both documents. An input that
  already carried tracked changes previously produced a package with two
  redlines layered on one another, where rejecting everything landed on a
  document neither side wrote; the base's own revisions are now resolved first,
  so the comparison is the only redline in the result and the round trip is
  exact.

- [#711](https://github.com/stella/folio/pull/711) [`e3a5f8f`](https://github.com/stella/folio/commit/e3a5f8f00253abe34ebf85b9bd3d8c1f05d22636) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Give the Arabic, narrow and ClearType faces documents actually request their own single-line height instead of the generic default.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - The AI-facing snapshot is linear in block count again. It resolved each block's
  position twice, and `ProseMirror`'s `resolve` scans a fragment from index 0, so
  a flat document cost O(blocks^2); the walk now carries the path it is already
  on. A 4,000-paragraph snapshot drops from 124ms to 24ms, and every reviewer
  read built on it drops with it. Fixing it surfaced a second defect: text in a
  table nested inside a hidden `w:trPr/w:hidden` row reached the snapshot,
  because the check consulted the nearest row rather than every enclosing row.

- [#707](https://github.com/stella/folio/pull/707) [`21be1d7`](https://github.com/stella/folio/commit/21be1d728e00131c519f2e5bd1187d97b04da318) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Place inline header and footer tables through the shared table placement rule, so `w:jc`, `w:tblInd` and `w:bidiVisual` behave as they do in the body.

- [#719](https://github.com/stella/folio/pull/719) [`745d509`](https://github.com/stella/folio/commit/745d5098a6dad02d9ef8c88d6a7871d042385a23) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Apply consecutive paragraph insertions at one document position in a single step.

- [#704](https://github.com/stella/folio/pull/704) [`cfefc7f`](https://github.com/stella/folio/commit/cfefc7f0c11d05244fbb6822aa3f1e5887a96a69) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A comparison covers more of what two documents differ by, and says it in the
  markup the format defines for it.

  A whole paragraph added or removed now carries its paragraph mark as well as its
  runs — `w:pPr/w:rPr/w:ins` and `w:pPr/w:rPr/w:del` — so accepting a deletion
  removes the paragraph instead of leaving a blank line, and rejecting an
  insertion closes the break instead of leaving an empty one.
  Resolving a deleted mark keeps the surviving paragraph's own properties.

  A list item that stopped being one is reported, and a paragraph inserted beside a
  list item is no longer silently made a further item of that list:
  `setBlockParagraphProperties` and the block insertions accept `listLevel: null`,
  which clears `w:numPr` the way `styleId: null` already clears `w:pStyle`.

  Additions past the base document's last block keep the target's order, so a
  paragraph and a table added after it no longer come out table first. Headers and
  footers pair by kind and document order when the two packages share no
  relationship id, so a comparison of two independently authored documents covers
  them instead of reporting them as present on one side only.

- [#696](https://github.com/stella/folio/pull/696) [`010e5c3`](https://github.com/stella/folio/commit/010e5c3c0e621a17ad4ca0b5f997f2e8c0785c36) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A tracked table-row insertion or deletion now marks the runs in its cells as
  well as the row, the way Word writes it, and resolves both halves together.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `compareDocx` returns the base package unchanged when it finds no difference,
  instead of re-serializing it. A 2,200-block comparison of two identical
  documents spent a second rewriting bytes nobody edited. A base that arrived
  carrying its own tracked changes is still serialized, because the compared base
  is its accepted view rather than the package as stored.

- [#702](https://github.com/stella/folio/pull/702) [`39894a3`](https://github.com/stella/folio/commit/39894a3c95530b671c0dc06709bf61573dbd5589) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Seven fidelity and review fixes:

  - A paragraph whose `w:pPr` states only `w:ilvl` now keeps the `w:numId` its
    style supplies, so a demoted styled list paragraph stays numbered.
  - Inserting a table row through a vertical merge extends the merge instead of
    splitting the grid.
  - Adjacent paragraphs in a table cell that share a border definition draw one
    frame with the `w:between` rule, as they already do elsewhere.
  - Every vertical `w:textDirection`, not only `btLr`, rotates its cell text.
  - An abrupt-closing HTML comment (`<!-->`) no longer swallows the rest of a
    paste.
  - A tracked replace whose two halves carry different timestamps is one review
    card again.
  - Striking a selection leaves another author's existing deletion attributed to
    them.

- [#695](https://github.com/stella/folio/pull/695) [`7356867`](https://github.com/stella/folio/commit/7356867bf96e9975ca08ae2c6d43830c1a97a54d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - A block inserted next to a block inside nested tables now lands at document
  level. `insertBeforeBlock`, `insertAfterBlock`, and `insertSignatureTable`
  escape the table their anchor sits in; they escaped only the innermost one, so
  an anchor two tables deep left the new block inside the outer cell. This is
  what made `compareDocx` refuse a pair whose base ends with a nested table and
  whose target appends a paragraph after it.

- [#727](https://github.com/stella/folio/pull/727) [`046302a`](https://github.com/stella/folio/commit/046302a9b1b0aef5d7a3dad8d3bf9c33e0f8babd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep physical tracked-change revision IDs unique when saving DOCX packages.

- [#707](https://github.com/stella/folio/pull/707) [`21be1d7`](https://github.com/stella/folio/commit/21be1d728e00131c519f2e5bd1187d97b04da318) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Measure `w:tblInd` from the leading cell's text edge in documents whose `compatibilityMode` predates the border-edge rule.

- [#724](https://github.com/stella/folio/pull/724) [`4562029`](https://github.com/stella/folio/commit/4562029cd050d4a44c8eb7bac63c796d04648b07) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Declare every ignorable namespace in serialized comment parts.

- [#726](https://github.com/stella/folio/pull/726) [`b74acb2`](https://github.com/stella/folio/commit/b74acb2ed0325a907f9967fb5c190daf0cf79ff6) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Serialize run, rendered-page-break, and table-cell properties in schema-valid OOXML order and form.
- Updated dependencies [[`3c5e627`](https://github.com/stella/folio/commit/3c5e627559b2cbd9d06e7c6dd7066488076d67b8), [`046302a`](https://github.com/stella/folio/commit/046302a9b1b0aef5d7a3dad8d3bf9c33e0f8babd), [`681923a`](https://github.com/stella/folio/commit/681923ab277b78acc69f3eeaea0262ec07fcf168)]:
  - @stll/docx-core@0.19.0

## 0.32.2

### Patch Changes

- [#691](https://github.com/stella/folio/pull/691) [`69299ce`](https://github.com/stella/folio/commit/69299cedd18d25cd11cdd689227caabf909586ab) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Split `insertAfterBlock` / `insertBeforeBlock` text on line breaks into consecutive paragraphs instead of one paragraph with embedded newlines, and report the split as a `splitMultilineText` normalization (`FolioAIEditApplyResult.normalizations`, surfaced through `suggest_changes`' `normalizations`). Only the first paragraph keeps `styleId` / `inheritFormatting`; later paragraphs get body formatting.

## 0.32.1

### Patch Changes

- [#689](https://github.com/stella/folio/pull/689) [`ab10444`](https://github.com/stella/folio/commit/ab104443c77afde1528148b813c6828db5a2f6e2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Parse legal-source drafts as GFM markdown plus `@` directives with `marked`: clause bodies, list items, and table cells keep inline emphasis, links, and code spans; markdown lists and pipe tables outside a directive become real list and table blocks. `compileMarkdownToContent` and `sanitizeExternalUrl` move into `@stll/docx-core`, and `@stll/folio-core`'s `fromMarkdown` now wraps them.
- Updated dependencies [[`ab10444`](https://github.com/stella/folio/commit/ab104443c77afde1528148b813c6828db5a2f6e2)]:
  - @stll/docx-core@0.18.0

## 0.32.0

### Minor Changes

- [#687](https://github.com/stella/folio/pull/687) [`267bfcd`](https://github.com/stella/folio/commit/267bfcd12a8c352c625f4c339df88acb67751859) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add a batch-level `precondition.documentVersion` to the document-operation contract, the `documentVersionMismatch` and `documentNotEditable` skip reasons with their recovery hints, a `queued` status and result list for host review-queue surfaces, and export `FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE`, the per-type property map the parser enforces.

### Patch Changes

- [#683](https://github.com/stella/folio/pull/683) [`63c2e75`](https://github.com/stella/folio/commit/63c2e75c79785f6b9117f6ded8753a27b588fb28) Thanks [@berticeek](https://github.com/berticeek)! - Read bounded DOCX archive entries through JSZip's platform-neutral internal stream instead of Node streams, so `ensureParaIds` and `loadDocxArchive` work in browsers and web workers.

## 0.31.2

### Patch Changes

- [#686](https://github.com/stella/folio/pull/686) [`f7fe238`](https://github.com/stella/folio/commit/f7fe2389e1efbc20a2bbf190f0ff59e05fd9c923) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve valid bookmark identities when creating bilingual DOCX copies.

- [#684](https://github.com/stella/folio/pull/684) [`3aa52ae`](https://github.com/stella/folio/commit/3aa52aee3b0b4e29564f15dd0b5b8a93bfe35a12) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep rewritten tracked-change clauses readable by matching words with their surrounding whitespace.

## 0.31.1

### Patch Changes

- [#679](https://github.com/stella/folio/pull/679) [`1fc03d2`](https://github.com/stella/folio/commit/1fc03d26df70e8d86fb6232b705e8be924943ee3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Size paragraph frames without an authored width to their content before resolving anchored alignment.

- [#681](https://github.com/stella/folio/pull/681) [`c11046d`](https://github.com/stella/folio/commit/c11046deece87684545045f99be24c4b550749d2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align collapsed borders between minimum-height table rows across page continuations.

## 0.31.0

### Minor Changes

- [#678](https://github.com/stella/folio/pull/678) [`8ad72ba`](https://github.com/stella/folio/commit/8ad72baa96ddabe1a106c7a2d9a4a08928fdd215) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add an opt-in stacked layout for bilingual source tables while preserving the inline layout as the default.

### Patch Changes

- [#676](https://github.com/stella/folio/pull/676) [`c87d755`](https://github.com/stella/folio/commit/c87d7556d88d71e8831041955404dfe774a2837f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep numbered-list markers visible when projecting paragraphs into bilingual columns.

## 0.30.0

### Minor Changes

- [#654](https://github.com/stella/folio/pull/654) [`faa8e0f`](https://github.com/stella/folio/commit/faa8e0f35148859cc8db9f061f30f5acaa60aec0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Route document I/O through the shared editor controller and enforce portable engine and projection boundaries.

- [#660](https://github.com/stella/folio/pull/660) [`9231301`](https://github.com/stella/folio/commit/9231301e004b5e9805fadaa5c4c259701e18898b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor document font alternate names during layout and rendering.

### Patch Changes

- [#673](https://github.com/stella/folio/pull/673) [`91e9a5d`](https://github.com/stella/folio/commit/91e9a5d1346053cbe95673c82d0208a3c803a086) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Position inline table borders from leading-edge indents while honoring inherited, bidirectional, and row-level justification.

- [#674](https://github.com/stella/folio/pull/674) [`3cae848`](https://github.com/stella/folio/commit/3cae84861078e3f61ac999cb9ef7b7cb27d9681f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep authored punctuation adjacent to generated footnote numbers.

- [#675](https://github.com/stella/folio/pull/675) [`934d425`](https://github.com/stella/folio/commit/934d425d721aecdb9fd20a5f9f6797d0650dae00) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve character spacing across word boundaries during line measurement.

- [#672](https://github.com/stella/folio/pull/672) [`f7c5205`](https://github.com/stella/folio/commit/f7c520579e79747055f79ef3794e2a5739b5a923) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align justified list continuation fitting with measured space capacity.

- [#668](https://github.com/stella/folio/pull/668) [`9129aa7`](https://github.com/stella/folio/commit/9129aa704458404fcf1e5f32d98d5bb67a2e9fb8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allow spaced percentage signs to wrap in non-East-Asian text.

- [#666](https://github.com/stella/folio/pull/666) [`33e9bf9`](https://github.com/stella/folio/commit/33e9bf9cc2816c454470f6721d05234fab5fe724) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor sparse legacy VML text-box inset values.

- [#664](https://github.com/stella/folio/pull/664) [`c7e2b00`](https://github.com/stella/folio/commit/c7e2b0047287630061ff9e559b36d55b7cb905b1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render bounded vector-only EMF logos as browser-safe SVG previews.

- [#670](https://github.com/stella/folio/pull/670) [`3329262`](https://github.com/stella/folio/commit/3329262ee9d3351e978a5c15d428a02f96bed4a7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep authored empty paragraphs as anchors for keep-with-next pagination.

- [#656](https://github.com/stella/folio/pull/656) [`948e678`](https://github.com/stella/folio/commit/948e678e95de2870451afdd899ad6f6939d134ea) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Guard invalid tab intervals and keep controller serialization free of Vue save side effects.

- [#663](https://github.com/stella/folio/pull/663) [`78d1138`](https://github.com/stella/folio/commit/78d1138f3ab3f222910d6d57b6ce352d917b2de7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align inline table indentation to the leading cell text edge.

- [#669](https://github.com/stella/folio/pull/669) [`5b8a0a3`](https://github.com/stella/folio/commit/5b8a0a3fad71a72c43c2b5127de610be401d2ed4) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve over-wide floating-table alignment, resolve page and text anchor frames with physical-page clamping, project wrap exclusions across columns and continuous sections with pre-block page advances, and default omitted section breaks to next-page transitions.

- [#658](https://github.com/stella/folio/pull/658) [`e347c0f`](https://github.com/stella/folio/commit/e347c0f2ab55407409d2cb1a4f7ad81510b44167) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Coalesce redundant hard breaks with pages already opened by section transitions.

- [#671](https://github.com/stella/folio/pull/671) [`d1b5972`](https://github.com/stella/folio/commit/d1b597218cfcebd4d3660408b8088909466f8d23) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize horizontal text scales across parsing, layout, and rendering.

- [#667](https://github.com/stella/folio/pull/667) [`a7c9d18`](https://github.com/stella/folio/commit/a7c9d185cd1593e6f34aaf4203f14452ed868d7c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve DrawingML vertical alignment and supported legacy text-box anchors through layout and save.

- [#659](https://github.com/stella/folio/pull/659) [`2330a00`](https://github.com/stella/folio/commit/2330a00ba952d7c1e71a3f6c1bc8c6fecf606683) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve section page numbers and default furniture when adjacent hard breaks share a physical page.

- [#661](https://github.com/stella/folio/pull/661) [`d39ed84`](https://github.com/stella/folio/commit/d39ed84600625c4bb3181618a2cbe14a8c00fe78) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve absolute VML image positioning without adding page artwork to text flow.

- [#662](https://github.com/stella/folio/pull/662) [`99826b6`](https://github.com/stella/folio/commit/99826b680a7d79cfa5a89e02dde45d501dc3b815) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Suppress automatic paragraph spacing at table-cell boundaries while preserving interior and authored spacing.

- [#665](https://github.com/stella/folio/pull/665) [`77fee80`](https://github.com/stella/folio/commit/77fee80ebb93192fb6b24a9a06e386bacb5eb72a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render bounded nested VML freeform paths with their authored local coordinates.
- Updated dependencies [[`a7c9d18`](https://github.com/stella/folio/commit/a7c9d185cd1593e6f34aaf4203f14452ed868d7c)]:
  - @stll/docx-core@0.17.3

## 0.29.0

### Minor Changes

- [#653](https://github.com/stella/folio/pull/653) [`9576702`](https://github.com/stella/folio/commit/957670247c26c04591c83430651717ba4df30417) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Resolve numbered REF fields from live bookmark and list state, and refresh safe cached results on save.

### Patch Changes

- [#645](https://github.com/stella/folio/pull/645) [`6d31019`](https://github.com/stella/folio/commit/6d310198e0197c276f93f5eab552ee8198ab3d80) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reuse equivalent paragraph measurements and preserve authored terminal blank-page spacing.

## 0.28.1

### Patch Changes

- [#648](https://github.com/stella/folio/pull/648) [`24d886a`](https://github.com/stella/folio/commit/24d886a32ee913830a3304b4dcf56e0c5f63c8af) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep earlier columns above footnote areas discovered during later-column layout.

- [#647](https://github.com/stella/folio/pull/647) [`865a8e4`](https://github.com/stella/folio/commit/865a8e4716392ed289c30c31601b785d7ceb4b72) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Advance continuous-section page restarts across shared-page content.

- [#650](https://github.com/stella/folio/pull/650) [`7ab1100`](https://github.com/stella/folio/commit/7ab1100aa103ee71f44fc298343f49fa7a115d5b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Apply level-aware OOXML style toggle inheritance while keeping direct run formatting explicit.

- [#646](https://github.com/stella/folio/pull/646) [`dba65db`](https://github.com/stella/folio/commit/dba65db5a6deaab14e54ee09d85500f6aec3e1bc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep short list markers from pulling text into a hanging indent.

## 0.28.0

### Minor Changes

- [#640](https://github.com/stella/folio/pull/640) [`73abf32`](https://github.com/stella/folio/commit/73abf32cb8c8738c52bf43a9293ccb167f9d0f76) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add server-side materialization of Folio Yjs snapshots into fidelity-preserving DOCX files.

### Patch Changes

- [#636](https://github.com/stella/folio/pull/636) [`33ab462`](https://github.com/stella/folio/commit/33ab462adf09196d1bef813cae924f262de16852) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound extracted DOCX table markdown to prevent resource exhaustion from padded rows.
- Updated dependencies [[`4582ad7`](https://github.com/stella/folio/commit/4582ad7671c31e757bbee0ae4d829186dd2be1bc)]:
  - @stll/docx-core@0.17.2

## 0.27.1

### Patch Changes

- [#641](https://github.com/stella/folio/pull/641) [`9050b4c`](https://github.com/stella/folio/commit/9050b4c688d5fbbf262d934b041ccce5c00463fe) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep bilingual columns and merged-cell widths valid, preserve signature-field rules, and maintain highlighted text contrast on dark document canvases.

- [#638](https://github.com/stella/folio/pull/638) [`cefef0a`](https://github.com/stella/folio/commit/cefef0a812d873138ee34026a6095b36b7d4dfc1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Escape theme-font, page-background and chapter-separator attributes on save;
  narrow theme-font references and `w:shd` colours at parse; drop the attached
  template reference from saved packages; convert only referenced media, with a
  package-wide decode budget; preflight XML resource limits at unzip and bound
  extracted text; bound `xmlns` declaration values; narrow run hyperlink targets
  in the painter; let the save path materialize a header/footer added through
  `createEmptyHeaderFooter`.

## 0.27.0

### Minor Changes

- [#634](https://github.com/stella/folio/pull/634) [`8bd6ea8`](https://github.com/stella/folio/commit/8bd6ea8c31b66e30143b8f28a050032c09ebeff1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose instance-scoped painted block geometry and layout-change subscriptions.

## 0.26.0

### Minor Changes

- [#632](https://github.com/stella/folio/pull/632) [`e1993d1`](https://github.com/stella/folio/commit/e1993d196734e1e1b69bc07759ec114a14a036b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Derive bilingual translation manifests from the canonical AI-edit snapshot, preserving structural-only paragraphs while rejecting unaddressable row handles.

### Patch Changes

- [#628](https://github.com/stella/folio/pull/628) [`693c320`](https://github.com/stella/folio/commit/693c32009623bd3e8aa98cbdb0e813b8cca55a64) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bound generated redline operations for fragmented formatting changes.

- [#630](https://github.com/stella/folio/pull/630) [`99d201f`](https://github.com/stella/folio/commit/99d201fc90341c13f090c7206b56613b52055f99) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Defer DOCX ZIP entry extraction until archive validation completes.

- [#629](https://github.com/stella/folio/pull/629) [`2bfc5c6`](https://github.com/stella/folio/commit/2bfc5c6cc8fde32b346b7a0f003527bc9d507631) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Escape relationship IDs when serializing OOXML attributes.

## 0.25.5

### Patch Changes

- [#624](https://github.com/stella/folio/pull/624) [`8724016`](https://github.com/stella/folio/commit/872401629e10ca7837bdcf59978d303ef7ccbac5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Parse, render, and round-trip adjusted right-brace shapes as editable geometry.

- [#625](https://github.com/stella/folio/pull/625) [`b10833d`](https://github.com/stella/folio/commit/b10833dedbca1682fdceaec5ddf8ab6f47aefc88) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Recognize localized table-of-contents styles so entry hyperlinks use TOC typography.

- [#623](https://github.com/stella/folio/pull/623) [`036b534`](https://github.com/stella/folio/commit/036b534d8c10d81ff084fce781498cb9e511b3b5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allow text editing when unsupported drawings are preserved as inert raw OOXML.
- Updated dependencies [[`8724016`](https://github.com/stella/folio/commit/872401629e10ca7837bdcf59978d303ef7ccbac5), [`036b534`](https://github.com/stella/folio/commit/036b534d8c10d81ff084fce781498cb9e511b3b5)]:
  - @stll/docx-core@0.17.1

## 0.25.4

### Patch Changes

- [#621](https://github.com/stella/folio/pull/621) [`d990d30`](https://github.com/stella/folio/commit/d990d30399f35518ef74bde05754f0168e8207c5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep paragraph measurement caching and RTL list geometry deterministic.

## 0.25.3

### Patch Changes

- [#619](https://github.com/stella/folio/pull/619) [`f176831`](https://github.com/stella/folio/commit/f176831c7be0a92a45b8937960c5b2d4bb21c254) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Match Word's RTL table, paragraph, and mixed-link geometry.

## 0.25.2

### Patch Changes

- [#611](https://github.com/stella/folio/pull/611) [`8400773`](https://github.com/stella/folio/commit/8400773cd915e20a2c9ee9ab9c3ed4049a483e10) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Persist resolved review views across every editable Word document story.

- [#617](https://github.com/stella/folio/pull/617) [`28da27c`](https://github.com/stella/folio/commit/28da27c1131375dcea56aed28155ff141267f864) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep resolved-view persistence stable with anchored and threaded comments across every document story.

- [#618](https://github.com/stella/folio/pull/618) [`610f7f9`](https://github.com/stella/folio/commit/610f7f920b1e324de65ef25148050adb461388aa) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep comment and reply IDs unique across story-scoped operations, with exhaustive resolved-view persistence coverage for comments in every editable story.

## 0.25.1

### Patch Changes

- [#614](https://github.com/stella/folio/pull/614) [`e17e8f5`](https://github.com/stella/folio/commit/e17e8f5ccbfd25141b9d20953821096df842beb2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve safe cached page boundaries inside split table rows.

## 0.25.0

### Minor Changes

- [#612](https://github.com/stella/folio/pull/612) [`75d34d2`](https://github.com/stella/folio/commit/75d34d291bd31e40f3c3c35fdbeb3c93529c0653) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Improve Arabic DOCX parity for no-wrap text boxes, boundary-owned section pagination, RTL tabs, and page-number formats.

### Patch Changes

- [#609](https://github.com/stella/folio/pull/609) [`21ef02a`](https://github.com/stella/folio/commit/21ef02a8454e0aa49aeeb11f2be4dd5cb51d18e2) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align Arabic header wrap bands, RTL table cells, and displayed URLs with Word layout.
- Updated dependencies [[`75d34d2`](https://github.com/stella/folio/commit/75d34d291bd31e40f3c3c35fdbeb3c93529c0653)]:
  - @stll/docx-core@0.17.0

## 0.24.0

### Minor Changes

- [#607](https://github.com/stella/folio/pull/607) [`069ba36`](https://github.com/stella/folio/commit/069ba36ac26a686d1b5cbbc2ee69c940cdb238f0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve independent complex-script typography on list markers.

### Patch Changes

- Updated dependencies [[`069ba36`](https://github.com/stella/folio/commit/069ba36ac26a686d1b5cbbc2ee69c940cdb238f0)]:
  - @stll/docx-core@0.16.0

## 0.23.1

### Patch Changes

- [#603](https://github.com/stella/folio/pull/603) [`d277e14`](https://github.com/stella/folio/commit/d277e14616d53028b40f906f15c75c68c93fad29) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Align unindented tables, preserve paragraph direction and section page numbering, and paginate independent table rows.

- [#605](https://github.com/stella/folio/pull/605) [`8f32b83`](https://github.com/stella/folio/commit/8f32b83d8006bc28bbc03083240a04dab1017c51) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve complex-script typography, Arabic list counters, and table direction.

## 0.23.0

### Minor Changes

- [#601](https://github.com/stella/folio/pull/601) [`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Avoid reparsing already validated agent document-operation batches while keeping core batch validation canonical.

### Patch Changes

- [#601](https://github.com/stella/folio/pull/601) [`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Use one bounded DOCX package validator with structured failure codes.

- [#601](https://github.com/stella/folio/pull/601) [`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve typed document values during auto-save recovery and validate SDT list items consistently.

- [#601](https://github.com/stella/folio/pull/601) [`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Type core command lookups and remove unchecked assertions from formatting, paragraph, and table wrappers.

- [#601](https://github.com/stella/folio/pull/601) [`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve complex-script theme fonts and consolidate typed selection formatting extraction.

- [#601](https://github.com/stella/folio/pull/601) [`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Validate nested paragraph property-change formatting at the ProseMirror boundary and consume it through typed paragraph attrs.
- Updated dependencies [[`6b636ad`](https://github.com/stella/folio/commit/6b636ad6a386fab1166c9a43441792258d352634)]:
  - @stll/docx-core@0.15.2

## 0.22.3

### Patch Changes

- [#597](https://github.com/stella/folio/pull/597) [`cfcaf36`](https://github.com/stella/folio/commit/cfcaf36927d858549f4b5ee39f4bb911312d59af) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Anchor body drawings to authored bottom margins and repaint when margin frames change.

- [#599](https://github.com/stella/folio/pull/599) [`d4e6d7f`](https://github.com/stella/folio/commit/d4e6d7ff18f0146a3fb0f97d3c427f8162df777d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve run font slots and underline color when formatting marks are rewritten at a selection or collapsed caret.

- [#590](https://github.com/stella/folio/pull/590) [`70c8bd4`](https://github.com/stella/folio/commit/70c8bd4f2f2fa5cfe36ed822688a43d22176c9fd) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Measure superscript and subscript text at the same scale used for painting.

- [#598](https://github.com/stella/folio/pull/598) [`cdde112`](https://github.com/stella/folio/commit/cdde112f1a85f7121a50b2250266093cfaddebae) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve direct paragraph spacing edits without inlining inherited line spacing, including explicit zero overrides.

- [#596](https://github.com/stella/folio/pull/596) [`550f0f8`](https://github.com/stella/folio/commit/550f0f8a064279ad572a7b16cfeba11ade3d023e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Honor the hanging-indent flag when generating paragraph CSS.

- [#595](https://github.com/stella/folio/pull/595) [`35608fc`](https://github.com/stella/folio/commit/35608fce912f6c7196dc86467cee3e0413e73355) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Parse every OOXML `ST_OnOff` spelling consistently across style and drawing attributes.

- [#600](https://github.com/stella/folio/pull/600) [`022b326`](https://github.com/stella/folio/commit/022b32636157d17601b1ec8a646e22ee4b68d485) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve typed errors when inspecting malformed XML package parts.

## 0.22.2

### Patch Changes

- [#587](https://github.com/stella/folio/pull/587) [`692e0ba`](https://github.com/stella/folio/commit/692e0ba01feabf7f852be6ecde053c7574b1a4d9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Upgrade `better-result` to v3.
- Updated dependencies [[`692e0ba`](https://github.com/stella/folio/commit/692e0ba01feabf7f852be6ecde053c7574b1a4d9)]:
  - @stll/docx-core@0.15.1

## 0.22.1

### Patch Changes

- [#584](https://github.com/stella/folio/pull/584) [`a7111a1`](https://github.com/stella/folio/commit/a7111a1eb558ce052c49062d045d28ba52130a84) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve accepted tracked changes, wrapped text, and tables when reading note stories.

## 0.22.0

### Minor Changes

- [#578](https://github.com/stella/folio/pull/578) [`fe0900d`](https://github.com/stella/folio/commit/fe0900df163f0317aed6218ba0b0e69a275eae44) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `createBilingualDocument` / `createBilingualDocx` to the server entry: lay a document body out as a two-column table (source | copy) with numbering and numbered paragraph styles cloned per language so both columns count independently. Repacking now appends numbering definitions and styles the model adds to the original parts.

## 0.21.0

### Minor Changes

- [#577](https://github.com/stella/folio/pull/577) [`9385d46`](https://github.com/stella/folio/commit/9385d464b85effe1d4aaa558ebea8e29cfcb84e3) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Headless DOCX report generation: `/server` exports typed builders (`heading`, `paragraph`, `run`, `table`, `pageBreak`, `hyperlink`, `bookmark`, `endnote`, `createTableOfContentsField`); external hyperlinks inside headers, footers, footnotes and endnotes now get relationships in their own rels part; `createEmptyDocument` initialises `package.relationships` so in-memory headers and footers materialise; `DocumentSettings.updateFields` round-trips; the Stella style set gains `Heading1`-`Heading6`, `TOCHeading`, `TOC1`-`TOC3`, `EndnoteReference` and `EndnoteText`; complex fields keep `w:dirty`/`w:fldLock` across parse and save.

### Patch Changes

- Updated dependencies [[`9385d46`](https://github.com/stella/folio/commit/9385d464b85effe1d4aaa558ebea8e29cfcb84e3)]:
  - @stll/docx-core@0.15.0

## 0.20.0

### Minor Changes

- [#574](https://github.com/stella/folio/pull/574) [`d6c389e`](https://github.com/stella/folio/commit/d6c389ef5082c58f1bb68c4b0d2a4d6364ddf79f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Remove the `documentKey` prop from the React and Vue `DocxEditor` (deprecated and ignored since the previous release) and the `useDocxEditor` option. Both adapters now hand the hidden-editor manager a per-load identity from their own loaders (`HiddenEditorManagerDeps.getDocumentIdentity`, required); the document-metadata fallback signature is gone.

## 0.19.0

### Minor Changes

- [#572](https://github.com/stella/folio/pull/572) [`b7b941a`](https://github.com/stella/folio/commit/b7b941af62a430217b86bba7b4c8c3e9a85d3ce1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reset the hidden editor on every document load, not only when the host's `documentKey` changes. `DocumentLoaderManager` now publishes a per-load identity (`setLoadedDocumentIdentity`) in the same commit as the history reset and orders parsed-document loads against in-flight buffer parses; the React `DocxEditor` feeds that identity to the paged editor, so loading a new `document`/`documentBuffer` (or `loadDocument`) into an editor with a live view replaces the painted content instead of leaving it on the previous document. The `documentKey` prop is deprecated and ignored.

### Patch Changes

- [#562](https://github.com/stella/folio/pull/562) [`418e5a1`](https://github.com/stella/folio/commit/418e5a1c18241dce37a36fbc128609ac72c3b161) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wait for the fonts the renderer actually uses, not just the names the document wrote.

  `resolveFontFamily` turns an authored family into a CSS stack that appends folio's bundled metric-compatible substitutes and a script fallback, so an authored `Arial` run paints its Arabic in the bundled Arabic face. The font-readiness gate collected only authored names, so it released the first layout before those faces had loaded; measurement taken against the pre-load fallback then disagreed with what was ultimately painted, by as much as a third of a line's width.

  The gate now expands each family through the resolver and waits for every concrete face in the stack. That also removes a hand-kept substitute table which duplicated the resolver's own mapping and was free to drift from it.

## 0.18.0

### Minor Changes

- [#568](https://github.com/stella/folio/pull/568) [`4197680`](https://github.com/stella/folio/commit/4197680777aee68ebcaf798ed28fa847463f429f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Generate tracked formatting redlines, enumerate every supported body revision, and expose a dedicated redline entry point.

### Patch Changes

- Updated dependencies [[`a38d902`](https://github.com/stella/folio/commit/a38d9025773e1bdba6bfbb4ffcab3dea1a943d6e)]:
  - @stll/docx-core@0.14.0

## 0.17.1

### Patch Changes

- [#566](https://github.com/stella/folio/pull/566) [`0e40ca0`](https://github.com/stella/folio/commit/0e40ca014bb241216dfce01dac0128a1fecbab58) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep justified tabbed first lines within the paragraph's right margin.

## 0.17.0

### Minor Changes

- [#565](https://github.com/stella/folio/pull/565) [`fdec7b5`](https://github.com/stella/folio/commit/fdec7b5e2662afe85c9d4df75ca87fd17cffaa61) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Expose source cell paragraphs alongside extracted DOCX table rows.

### Patch Changes

- [#560](https://github.com/stella/folio/pull/560) [`b9b7d25`](https://github.com/stella/folio/commit/b9b7d25f9e4da20d0225eb95a48614b07dcc85da) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render complex-script text in the font the document asks for.

  Word resolves a run's font per character across three slots: `w:eastAsia` for CJK, `w:cs` for Arabic, Hebrew, Indic and South-East Asian text, and `w:ascii`/`w:hAnsi` for everything else. folio honoured the first and third but never the second: `w:cs` was parsed and round-tripped yet never reached layout, so a document written the standard way (`w:ascii="Calibri"` with `w:cs="Traditional Arabic"`) measured and painted its Arabic in Calibri.

  The complex-script slot now flows through the bridge, the measurer and the painter on the same path the East-Asian slot already used, so the two stay segmented identically and line wrapping keeps matching what is drawn. Script segmentation now reports which of the three slots a character selects rather than a CJK yes/no.

- [#556](https://github.com/stella/folio/pull/556) [`768baab`](https://github.com/stella/folio/commit/768baab2a609384057e51bea45957bea9df1bbe1) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep cursive words joined when a run boundary splits them mid-word.

  Browsers shape across an inline box boundary only while no shaping-relevant property changes. A bold, italic, resized, or refaced run inside an Arabic word therefore stopped shaping at that boundary, and the word rendered as isolated letter forms. Word joins straight through the same boundary.

  The painter now emits a zero-width joiner on each side of such a boundary. Each joiner sits in its own span, so run text nodes keep their exact ProseMirror offsets. Colour and underline boundaries, which is what tracked changes and comment anchors use, already shaped correctly and are left untouched.

  Joining classes are generated from the pinned Unicode Character Database rather than hand-listed, so combining marks are classified correctly.

- [#558](https://github.com/stella/folio/pull/558) [`14eee8f`](https://github.com/stella/folio/commit/14eee8f644694d3a47060f5445df1465cf91ae77) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Wait for the complex-script and East-Asian fonts before the first layout, not just the Latin ones.

  The font-readiness gate collected only the `ascii` and `hAnsi` slots of a run's font. Word writes the Arabic and Hebrew face into `w:cs` and the CJK face into `w:eastAsia`, so a document that styles those scripts the standard way had its font ignored by the gate. The first layout then measured a fallback face while the painter later drew the real one once it loaded, and the two disagreed for exactly the scripts whose advances differ most from a Latin fallback.

## 0.16.0

### Minor Changes

- [#552](https://github.com/stella/folio/pull/552) [`198e98e`](https://github.com/stella/folio/commit/198e98e8dd7fd1c2b76eff944f92602cd3c1bdfb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Extract DOCX tables as markdown rows instead of a flat row-major paragraph list, so a cell stays associated with its column. `ExtractedDocxParagraph` gains an optional `tableRow` describing which table a row belongs to and whether it carries cells.

### Patch Changes

- [#551](https://github.com/stella/folio/pull/551) [`1670559`](https://github.com/stella/folio/commit/16705596439f650174891da6f9e29fe97b0e2e39) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Seed headless `paraId` attributes by rebuilding the document instead of applying one ProseMirror step per paragraph, removing two quadratics from `FolioDocxReviewer.fromBuffer` (-29% on a real DOCX corpus, -45% on the largest file). Also seeds RTL base direction explicitly, which previously rode on the paraId transaction and was skipped entirely for documents that already carried Word-authored paraIds.

- [#553](https://github.com/stella/folio/pull/553) [`e5bd094`](https://github.com/stella/folio/commit/e5bd094c34405a2ec32ba861448d5e35629ed085) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Bump `dompurify` to 3.4.13, which fixes an XSS advisory where an `IN_PLACE` hook removal left a detached subtree executable.

## 0.15.13

### Patch Changes

- [#545](https://github.com/stella/folio/pull/545) [`2bb8980`](https://github.com/stella/folio/commit/2bb8980f7d8a3e5db5bc1863ea5bca0a8316dc10) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Clear inherited highlight and shading across AI-replaced text while preserving suggestion rejection.

- [#545](https://github.com/stella/folio/pull/545) [`2bb8980`](https://github.com/stella/folio/commit/2bb8980f7d8a3e5db5bc1863ea5bca0a8316dc10) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Restart a nested list counter after returning to its parent level, and never
  paint a non-finite counter value into a document marker.

- [#545](https://github.com/stella/folio/pull/545) [`2bb8980`](https://github.com/stella/folio/commit/2bb8980f7d8a3e5db5bc1863ea5bca0a8316dc10) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prefer collapsed line-edge spans when adjacent runs share a caret boundary, and
  advance the caret by the measured width of each collapsed trailing space.

## 0.15.12

### Patch Changes

- [#541](https://github.com/stella/folio/pull/541) [`b00cb19`](https://github.com/stella/folio/commit/b00cb19b12e3e56885d5eef5b55f63c488e15596) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Anchor CJK IME candidate windows to the painted document caret.

- [#544](https://github.com/stella/folio/pull/544) [`f85e926`](https://github.com/stella/folio/commit/f85e926599e87ca3eab4966eff06aa1e45882643) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep the caret aligned with its text row after typing collapsible whitespace.

## 0.15.11

### Patch Changes

- [#538](https://github.com/stella/folio/pull/538) [`ccb1953`](https://github.com/stella/folio/commit/ccb19531e3a0780573d6c7c7a83842434b138b56) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Create required DOCX package parts when adding footnotes or endnotes to a document that had none.

## 0.15.10

### Patch Changes

- Updated dependencies [[`886b6f6`](https://github.com/stella/folio/commit/886b6f6cd0f2a407c872c90c2ef294192ea3bc0c)]:
  - @stll/docx-core@0.13.0

## 0.15.9

### Patch Changes

- Updated dependencies [[`2ef7445`](https://github.com/stella/folio/commit/2ef74452c34153a09b1af7496f27d8abd2074efc), [`51cbe7d`](https://github.com/stella/folio/commit/51cbe7d708fe44786a7fe8165c42225d2e35b93e)]:
  - @stll/docx-core@0.12.0

## 0.15.8

### Patch Changes

- Updated dependencies [[`ae98003`](https://github.com/stella/folio/commit/ae98003cf69b31ce86b44b5be0ff30834ee71455)]:
  - @stll/docx-core@0.11.0

## 0.15.7

### Patch Changes

- Updated dependencies [[`b4c2b15`](https://github.com/stella/folio/commit/b4c2b1536f1e3483ed0fa55e72c564d543964a41)]:
  - @stll/docx-core@0.10.0

## 0.15.6

### Patch Changes

- Updated dependencies [[`d0ad1db`](https://github.com/stella/folio/commit/d0ad1db29b9fc6758c77512eeba6b093d539c3b3)]:
  - @stll/docx-core@0.9.0

## 0.15.5

### Patch Changes

- [#522](https://github.com/stella/folio/pull/522) [`003ea02`](https://github.com/stella/folio/commit/003ea02f4ac98c29529da7fc4a92bef9d06d9c63) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Render OOXML symbol characters with their declared fonts and preserve them through editing.

- Updated dependencies [[`ff873db`](https://github.com/stella/folio/commit/ff873db582719ecf692391bb90054daf25cb0adc), [`003ea02`](https://github.com/stella/folio/commit/003ea02f4ac98c29529da7fc4a92bef9d06d9c63)]:
  - @stll/docx-core@0.8.0

## 0.15.4

### Patch Changes

- Updated dependencies [[`e40ccd4`](https://github.com/stella/folio/commit/e40ccd435ea17102fc75910b175e2f78561ac359)]:
  - @stll/docx-core@0.7.0

## 0.15.3

### Patch Changes

- [#517](https://github.com/stella/folio/pull/517) [`4c943b3`](https://github.com/stella/folio/commit/4c943b308ed29f4ee76226ccf142f60477534080) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve renderable drawing namespaces and mixed-section page order during DOCX round-trips.

- [#511](https://github.com/stella/folio/pull/511) [`1e5681b`](https://github.com/stella/folio/commit/1e5681bfaa71082bfbf6d795a004a14ca76ed61b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize omitted offsets on positioned legacy text boxes.

- Updated dependencies [[`5ea99cd`](https://github.com/stella/folio/commit/5ea99cd32d44db94afac8bb44c31c8f32bc2aa19), [`b003927`](https://github.com/stella/folio/commit/b003927467052dcc6c6c2c3ddd66cebf057e7f84)]:
  - @stll/docx-core@0.6.0

## 0.15.2

### Patch Changes

- [#504](https://github.com/stella/folio/pull/504) [`adc62f3`](https://github.com/stella/folio/commit/adc62f3b1a2357a22b431ba715bd46a11533e9b7) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit unequal-width section column settings without companion attributes.

- [#497](https://github.com/stella/folio/pull/497) [`3b28632`](https://github.com/stella/folio/commit/3b28632025d3798b4bd4b9c8268fbb444237ce6c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve Word positional tabs across parsing, editing, and serialization.

- [#502](https://github.com/stella/folio/pull/502) [`8b2535e`](https://github.com/stella/folio/commit/8b2535ec268eba128895251a3d573af29c241b16) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve localized comment annotation-reference formatting across DOCX serialization.

- [#498](https://github.com/stella/folio/pull/498) [`c915317`](https://github.com/stella/folio/commit/c91531724c50c2199f679019846f5db0daeafe60) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize nested hyperlink runs even when the hyperlink is the paragraph's only child.

- [#501](https://github.com/stella/folio/pull/501) [`cc17b71`](https://github.com/stella/folio/commit/cc17b7140a5cbf943b338a2242237ac1e243ee95) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit complex-script formatting overrides during DOCX serialization.

- [#499](https://github.com/stella/folio/pull/499) [`18a900f`](https://github.com/stella/folio/commit/18a900f88e1cb313a6747b006e9431b4a020e499) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Assign deterministic, collision-free IDs to parsed drawings that omit them.

- [#493](https://github.com/stella/folio/pull/493) [`b06a26d`](https://github.com/stella/folio/commit/b06a26d6efeb818d12ed78799a626f5d058494e8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve DrawingML gradient fill details across text boxes and other shared fill consumers.

- [#496](https://github.com/stella/folio/pull/496) [`d762d06`](https://github.com/stella/folio/commit/d762d061fb0693d1b9d81803ba6fff9227de048d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve inline text box wrapping across document saves.

- [#506](https://github.com/stella/folio/pull/506) [`55a5913`](https://github.com/stella/folio/commit/55a5913202a82e1eee9caef2b59deab8be7e6a00) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicitly empty initials on comments.

- [#490](https://github.com/stella/folio/pull/490) [`ea689e2`](https://github.com/stella/folio/commit/ea689e24e1c7744715bee672554167e952e0d02b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit no-highlight run formatting when saving DOCX files.

- [#494](https://github.com/stella/folio/pull/494) [`0a613b8`](https://github.com/stella/folio/commit/0a613b879914f40ea2fc143caaed54fa2bd9412e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored DrawingML outline details across shapes and text boxes.

- [#500](https://github.com/stella/folio/pull/500) [`7a0c72f`](https://github.com/stella/folio/commit/7a0c72f32aa19b0f3956ed02aff0c2adaf23c8b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize omitted floating drawing positions to deterministic zero offsets.

- [#492](https://github.com/stella/folio/pull/492) [`e254b3b`](https://github.com/stella/folio/commit/e254b3bd681176138e757225025512f050038224) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize syntactically empty Word runs without hiding unsupported run payloads.

- [#505](https://github.com/stella/folio/pull/505) [`4ffed02`](https://github.com/stella/folio/commit/4ffed0279600c89db755d3aa385e945f7dbdd88a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve authored clear behavior on every break type.

- [#503](https://github.com/stella/folio/pull/503) [`fe20374`](https://github.com/stella/folio/commit/fe20374491d851dec25bb9754526d72f8cfdd561) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicitly disabled first-page section headers and footers.

- [#495](https://github.com/stella/folio/pull/495) [`4b3cdc8`](https://github.com/stella/folio/commit/4b3cdc8b3b9a103866107c20f34882e7196f7d39) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize no-op zero-height table row rules.

- [#507](https://github.com/stella/folio/pull/507) [`008fd68`](https://github.com/stella/folio/commit/008fd68ab8c3686a575f0dfdefcd9ba2ac77aa5e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit disabled emphasis marks in run formatting.

- Updated dependencies [[`3b28632`](https://github.com/stella/folio/commit/3b28632025d3798b4bd4b9c8268fbb444237ce6c), [`8b2535e`](https://github.com/stella/folio/commit/8b2535ec268eba128895251a3d573af29c241b16), [`b06a26d`](https://github.com/stella/folio/commit/b06a26d6efeb818d12ed78799a626f5d058494e8), [`0a613b8`](https://github.com/stella/folio/commit/0a613b879914f40ea2fc143caaed54fa2bd9412e)]:
  - @stll/docx-core@0.5.2

## 0.15.1

### Patch Changes

- [#440](https://github.com/stella/folio/pull/440) [`8df8081`](https://github.com/stella/folio/commit/8df80815d500c213680b273ce4ad98f19f950d8b) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Parse theme-colored straight connectors as editable shapes.

- [#481](https://github.com/stella/folio/pull/481) [`dafe214`](https://github.com/stella/folio/commit/dafe21438306ed4681c4c2b7c3a47478523f2646) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep cached rendered page-break metadata stable when later inline content is omitted.

- [#464](https://github.com/stella/folio/pull/464) [`d64e659`](https://github.com/stella/folio/commit/d64e65904cf8e407f5f36157790755c582cfe18e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Avoid creating unused image parts for raw OOXML drawing previews.

- [#470](https://github.com/stella/folio/pull/470) [`db702de`](https://github.com/stella/folio/commit/db702dea23ca7f7374031c13aac64ad2f520d981) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve Word font script hints across document edits and saves.

- [#472](https://github.com/stella/folio/pull/472) [`76c46d5`](https://github.com/stella/folio/commit/76c46d53f980c1651c83bd600bfbd565822cfa53) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve distinct image names, descriptions, and titles when saving DOCX files.

- [#453](https://github.com/stella/folio/pull/453) [`9606a11`](https://github.com/stella/folio/commit/9606a115a39bc9d185a64f451386a07b342ccad0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit portrait section orientation when saving DOCX files.

- [#442](https://github.com/stella/folio/pull/442) [`4088787`](https://github.com/stella/folio/commit/408878732db508d72d269d8da3704278b97f500a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Speed up DOCX serialization by reusing unchanged package entries.

- [#452](https://github.com/stella/folio/pull/452) [`c36a6cd`](https://github.com/stella/folio/commit/c36a6cd73cf5d568c7a31b6aea60a0ff79de2637) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit zero first-line paragraph indents during DOCX saves.

- [#449](https://github.com/stella/folio/pull/449) [`987c956`](https://github.com/stella/folio/commit/987c9564faa048f2ed3162e4903001b2afbd3650) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep whitespace-sensitive text models stable across save and reopen.

- [#475](https://github.com/stella/folio/pull/475) [`7f5ab65`](https://github.com/stella/folio/commit/7f5ab655e532a840733677687749ee7a86dfcced) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep preserved OOXML fragments stable across save and reopen.

- [#458](https://github.com/stella/folio/pull/458) [`0528c1c`](https://github.com/stella/folio/commit/0528c1ca6733df2bc6a20ce29a485b0a9ec564bb) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicitly authored solid shape outlines when saving DOCX files.

- [#456](https://github.com/stella/folio/pull/456) [`5152140`](https://github.com/stella/folio/commit/5152140be45bef9ceebe0939ca99bb745b125dec) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep captured VML run XML stable across repeated document round trips.

- [#443](https://github.com/stella/folio/pull/443) [`5c42d4e`](https://github.com/stella/folio/commit/5c42d4e77934fe995f52d954eb495022fe041e7e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Speed up DOCX parsing with batched package extraction and lower-allocation XML lookups.

- [#480](https://github.com/stella/folio/pull/480) [`c66ecc7`](https://github.com/stella/folio/commit/c66ecc733f52b9d7ffd19d4f28c87a02176fb6d0) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve omitted image wrap distances when saving documents.

- [#461](https://github.com/stella/folio/pull/461) [`9ce81df`](https://github.com/stella/folio/commit/9ce81dffee6ed812b91ee7a5cdb839c5d2d0f690) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve clickable image hyperlinks across DOCX save and reopen.

- [#460](https://github.com/stella/folio/pull/460) [`2f85509`](https://github.com/stella/folio/commit/2f85509b0fbf9e69e043991f32805ea90c48e033) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep untouched picture watermark headers byte-stable when documents are saved.

- [#473](https://github.com/stella/folio/pull/473) [`352c66a`](https://github.com/stella/folio/commit/352c66a8784ab525f70cca01d9148f3a0a57da42) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve empty comment package parts during no-op document saves.

- [#468](https://github.com/stella/folio/pull/468) [`142a301`](https://github.com/stella/folio/commit/142a3017846535748588b2f79ca0ce4f7c226247) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit left-to-right section bidi settings across DOCX save and reopen.

- [#469](https://github.com/stella/folio/pull/469) [`862b62b`](https://github.com/stella/folio/commit/862b62b74ce31ac2d7a27abda4c4cec2c168df1a) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve absent shape wrap distances when saving DrawingML.

- [#465](https://github.com/stella/folio/pull/465) [`0db58c8`](https://github.com/stella/folio/commit/0db58c84dc713e0c86840af5f4a56bd78d3c2383) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Allow DOCX parsing to omit rowless placeholder tables.

- [#474](https://github.com/stella/folio/pull/474) [`8b532b8`](https://github.com/stella/folio/commit/8b532b84594554323f5bee96a9ec6f31d8540c94) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Skip painter text measurement for lines that do not need tab or horizontal-scale geometry.

- [#459](https://github.com/stella/folio/pull/459) [`f124920`](https://github.com/stella/folio/commit/f12492029ef25eeedeb67c87dae639c1fe097bd9) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicitly authored empty shape names when saving DOCX files.

- [#446](https://github.com/stella/folio/pull/446) [`3c6d290`](https://github.com/stella/folio/commit/3c6d290c6c7478fdebaae16b9bbd0016e09b6d07) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Parse document XML with a faster single-pass path and compatibility fallback.

- [#483](https://github.com/stella/folio/pull/483) [`3fdbea2`](https://github.com/stella/folio/commit/3fdbea22fec5046273209c39516db9648d692a3c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce paragraph layout text measurements by reusing visible word widths for ordinary trailing whitespace.

- [#444](https://github.com/stella/folio/pull/444) [`580d7bf`](https://github.com/stella/folio/commit/580d7bff9952a2dd5d7e32aa98e8149decbc2414) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Reduce allocation overhead in namespace-tolerant XML name lookups.

- [#477](https://github.com/stella/folio/pull/477) [`fd04f0b`](https://github.com/stella/folio/commit/fd04f0bb8fe624f618dd10f46abf29e5240378c5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve paragraph and run formatting when serializing comment content.

- [#479](https://github.com/stella/folio/pull/479) [`4eda1fb`](https://github.com/stella/folio/commit/4eda1fb28a18cdfc536bdace1a3c23337aaded41) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Keep complex fields structurally stable across repeated DOCX save and reopen cycles.

- [#466](https://github.com/stella/folio/pull/466) [`a17c937`](https://github.com/stella/folio/commit/a17c9374fd44138487b36e2010dc44af9696ebfe) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve the authored simple-field structure when saving DOCX files.

- [#455](https://github.com/stella/folio/pull/455) [`0e959dc`](https://github.com/stella/folio/commit/0e959dcf1c331ff4a65b40a47894670777c38307) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Ignore negative OOXML list levels instead of rejecting the document model.

- [#457](https://github.com/stella/folio/pull/457) [`316d1dc`](https://github.com/stella/folio/commit/316d1dc92b11cfd46c55c12e603eba656c8f973d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Normalize the default paragraph tab leader during DOCX parsing.

- [#476](https://github.com/stella/folio/pull/476) [`b166291`](https://github.com/stella/folio/commit/b166291647e19bfe39697ca665c58868f6c1a17d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Prevent compatibility text boxes from being duplicated during document saves.

- [#467](https://github.com/stella/folio/pull/467) [`14e508f`](https://github.com/stella/folio/commit/14e508fb8960d306a8691e9097825e2eda35eb90) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve DrawingML background and text theme colors when saving shapes.

- [#462](https://github.com/stella/folio/pull/462) [`263a646`](https://github.com/stella/folio/commit/263a646e9c5024e9353a2ab1ab5f03be8d16cc91) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve conditional table-row formatting across DOCX save and reopen.

- [#454](https://github.com/stella/folio/pull/454) [`b0b0fdf`](https://github.com/stella/folio/commit/b0b0fdf4f12891c6b459d262199b53a58da32b04) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve explicit single-column section counts when saving DOCX files.

- [#478](https://github.com/stella/folio/pull/478) [`a285d18`](https://github.com/stella/folio/commit/a285d18f3120efae15ba7866fa8e5c9d7372d7e5) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve an absent dropdown selection when saving untouched content controls.

- [#471](https://github.com/stella/folio/pull/471) [`2556c9b`](https://github.com/stella/folio/commit/2556c9b25a3c8b78f2ae02f53520017a35c75c8c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve cached rendered-page starts before compatibility-wrapped paragraph content.

- [#450](https://github.com/stella/folio/pull/450) [`4f446fb`](https://github.com/stella/folio/commit/4f446fb3acd70d28a9457a2d17b0811e37b427cc) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Save URI-encoded SVG image data URLs without requiring base64 encoding.

- [#463](https://github.com/stella/folio/pull/463) [`c6abe66`](https://github.com/stella/folio/commit/c6abe6678c1eadf4fa21aa719243668c7dc7391f) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Preserve the distinction between absent and explicitly disabled shape fills.

- Updated dependencies [[`db702de`](https://github.com/stella/folio/commit/db702dea23ca7f7374031c13aac64ad2f520d981), [`76c46d5`](https://github.com/stella/folio/commit/76c46d53f980c1651c83bd600bfbd565822cfa53), [`9ce81df`](https://github.com/stella/folio/commit/9ce81dffee6ed812b91ee7a5cdb839c5d2d0f690)]:
  - @stll/docx-core@0.5.1

## 0.15.0

### Minor Changes

- [#437](https://github.com/stella/folio/pull/437) [`a630992`](https://github.com/stella/folio/commit/a6309920b87ba9db64a15e435bafcb83ece51a33) Thanks [@jan-kubica](https://github.com/jan-kubica)! - `fromMarkdown` now synthesizes `document.package.numbering` for the ordered/bullet
  lists it emits, so `createDocx(fromMarkdown(markdown))` no longer throws
  `DocxModelValidationError: Numbering definition N is missing` and round-trips
  through `docxToMarkdown` unchanged.

  Added `mergeDocumentContent(target, source)`, a general helper for appending one
  document's content onto another. It renumbers any `numId`/`abstractNumId` the
  source carries to sit above the target's existing numbering range, so merging
  `fromMarkdown`'s output into a styled preset (e.g.
  `createStellaStyleDocumentPreset()`) can no longer collide with numbering the
  preset already reserves — previously a markdown list could silently render with
  the preset's own clause/definition numbering instead of a plain bullet/number.

## 0.14.1

### Patch Changes

- [#434](https://github.com/stella/folio/pull/434) [`70217a6`](https://github.com/stella/folio/commit/70217a6fabaf0807807f2c790da856fdd9108c5d) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Correct the `@stll/docx-core` dependency range: folio-core imports `normalizeRevisionId`, which only exists in docx-core 0.5.0, but 0.14.0 was published declaring `^0.4.0` (excludes 0.5.0 under 0.x semver). Republish so the range resolves to `^0.5.0`.

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
