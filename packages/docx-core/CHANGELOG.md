# @stll/docx-core

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
