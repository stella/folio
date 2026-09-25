# @stll/docx-core

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
