# The container contract

folio loses things quietly. A parser walks a container's children, recognises
the ones it models, and lets the rest fall off the end of a `switch`; a
serializer rebuilds the container from the model and writes back only what the
model holds. Neither step reports anything, so the loss shows up as a Word
repair dialog, a reviewer's revision that vanished, or nothing at all.

The contract makes the loss a decision somebody wrote down, and the survival law
checks that decision by running it.

## The two halves

| Half             | Where                                | What it guarantees                                                                                                           |
| ---------------- | ------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------- |
| The survival law | `scripts/lib/container-survival/`    | Every pair the schema allows is exercised against a real save, and every loss is recorded with the mechanism that caused it. |
| The contract     | `specifications/container-contract/` | Every one of those pairs has a recorded disposition, and a `dropped` one names why and what would change it.                 |

They check each other. A pair the contract calls `modelled` or
`captured-verbatim` that the law shows lost fails CI. A pair the contract calls
`dropped` that the law shows surviving fails CI too — otherwise a fix would
leave a stale confession behind and the list would only ever grow. Both
directions run on every pull request, in the `Container contract and survival
census` job of `.github/workflows/ci.yml`; the job is where "fails CI" in this
document is cashed out.

## The survival law

The law's universe is derived, not listed. `schemaSpace.ts` walks the committed
schema graph (`specifications/generated/docx-transitional-schema.gen.json`, no
network) from the roots of the parts folio rebuilds on a save: `w:document`,
`w:styles`, `w:numbering`, `w:settings`, `w:fonts`, `w:webSettings`,
`w:comments`, `w:footnotes`, `w:endnotes`, `w:hdr` and `w:ftr`. The roots and
the packaging live in one table (`REBUILT_PARTS`), so a root cannot be walked
without saying which part carries it. It yields two kinds of pair:

- **(container, allowed child)** — every element the schema lets a container hold;
- **(element, allowed attribute)** — every attribute the schema lets an element carry.

For each, `fixture.ts` synthesises a minimal package: the cheapest chain of
elements from a rebuilt part's root down to the container, each level carrying
the attributes its type requires and the siblings its content model requires,
with the subject at the ordinal its particle declares. The part is written at
its own path with the content-type override and relationship that make a reader
find it, and a `w:hdr` or `w:ftr` fixture gets the section reference without
which nothing opens it. A fixture that does not itself validate is reported as
**unrepresentable** and not run, because a law that fails on the generator's own
invalid markup proves nothing.

Then four laws run, reported separately because they fail for different reasons
and are fixed in different places:

- **L1 parse** — `parseDocx` does not throw.
- **L2 serialize** — parse, remove the verbatim captures that replay would hand
  back, save, and find the subject in the saved part with an equal value.
- **L3 editor** — the same through `toProseDoc`/`fromProseDoc`.
- **L4 schema** — the part L2 wrote carries no new schema violation.

L2 is the one with teeth. folio replays captured bytes whenever a fingerprint
says the model still agrees with them, so a round trip over an untouched
document exercises the capture machinery rather than the serializers. The
stripping is the corpus `reserialize` invariant's own forcing mechanism, imported
from `scripts/lib/corpus-invariants/reserialize.ts` rather than restated.

### Loss mechanisms

Each names a different place to fix it:

| Mechanism                             | Means                                                                                      |
| ------------------------------------- | ------------------------------------------------------------------------------------------ |
| `the-container-itself-is-lost`        | The pair went with its container; fix the container and they all follow.                   |
| `never-parsed`                        | No parser reads it and no capture keeps it.                                                |
| `parsed-but-not-serialized`           | The model holds it and no serializer writes it.                                            |
| `serialized-only-via-verbatim-replay` | It survives an untouched save and not an edited one.                                       |
| `replay-rejected`                     | A capture holds it and a gate refuses the capture, forcing a rebuild that cannot write it. |
| `lost-in-the-editor-projection`       | It survives a save but not the ProseMirror round trip.                                     |
| `present-with-a-different-value`      | It comes back respelled.                                                                   |

### What counts as an equal value

Three re-spellings are folio's design rather than its defects, and the law asks
folio's own tables rather than restating them, so the two cannot drift:

- `ST_OnOff` has six spellings of two values, and folio canonicalises them.
- A Strict measure or percentage is re-spelled as Transitional, because folio
  rebuilds every package as Transitional. The law converts with
  `transitionalSlotEncoding`, the generated table `captureVerbatimXml` uses.
- A revision element's `w:id` is a physical wrapper id, re-minted on every save
  so ids stay unique across a package. The elements that carry one come from
  `REVISION_ELEMENT_NAMES` in `revisionIdNormalization.ts`.

Anything else that comes back different is `present-with-a-different-value`,
which is a finding.

### Fixture realism

Two small tables in the generator decide what a fixture looks like, and neither
decides anything the contract decides:

- `DETOUR_ELEMENTS` and `BLOCK_CONTAINERS`/`BLOCK_CHILDREN` in `schemaSpace.ts`
  cost a step through a transparent wrapper more than a step down the
  structural spine, and the two costs stack. The schema lets a `w:body` hold an
  `m:oMath` directly and lets a run-level `w:ins` sit straight under the body;
  documents put equations in paragraphs and tracked changes on runs. Without
  the weighting, every maths pair measures how folio treats a bare `m:oMath` in
  a body, and every `CT_RunTrackChange` pair measures a tracked insertion with
  no paragraph — a whole block of pairs was charged to the wrong container for
  exactly that reason.
- `SEED_CHILDREN` in `fixture.ts` gives a container the content it needs to
  survive at all — a row in a table, a paragraph in a cell, a numbering
  reference in a `w:numPr`, a run in any of the four transparent inline
  wrappers, a run in each of the four revision wrappers. A container folio
  prunes for being empty would report every pair inside it as lost. Two kinds
  of container are pruned for that reason and neither is a defect: a
  transparent wrapper the editor carries as a mark on its content, and a
  revision, which is the same shape one level up — `w:ins` is an insertion mark
  over the inline nodes it holds, so an empty one has no leaf to carry it.
  Text seeded into a `w:del` or a `w:moveFrom` is `w:delText`, which is what
  Word writes there; a `w:t` would measure folio's own correction instead.
  A seed is chosen by the container's content model, not by its name: every
  `w:sdt` declares a `w:sdtContent`, but that content is paragraphs under a
  block content control, runs under an inline one, rows inside a table and
  cells inside a row. Seeding by the name alone wrote block content into a
  run-level control, and the pairs under it were counted unrepresentable
  instead of measured.

### What is skipped, and why

- **Parts a repack replays verbatim.** Removing the capture slots makes the
  _element_ serializers run; it does not make a _part_ serializer run. A repack
  copies `word/styles.xml`, `word/numbering.xml`, `word/settings.xml`,
  `word/fontTable.xml`, `word/webSettings.xml`, `word/footnotes.xml` and
  `word/endnotes.xml` through byte for byte, so the forced leg hands back the
  fixture unchanged and every pair in them would read as surviving on the
  strength of a file copy. Recording that as `modelled` would put a disposition
  on a slot no model holds, so the law compares the bytes and reports the pair
  unrepresentable with the part named. The test is the bytes rather than a list
  of parts, so a part folio starts rebuilding starts being measured with no
  change to the law. Closing this needs the forcing to reach part level — the
  same idea one layer up — and it is where the styles-part defects live.
- **Content models the builder cannot satisfy mechanically.** A generated
  fixture that fails the schema validator is counted as unrepresentable, with
  the violation that made it so, rather than as a passing pair.

Both are counted and printed by reason; neither is silently treated as passing.

## The contract

`specifications/container-contract/contract.json` holds one entry per pair:

```json
"{…}tblGrid|{…}CT_TblGrid/{…}tblGridChange": { "disposition": "captured-verbatim" }
"{…}p|{…}CT_P@{…}rsidR": { "disposition": "dropped", "reason": "neverParsed" }
```

`modelled` means a parser reads it into the typed model and a serializer writes
it back: it survives an edit, the editor, and a rebuild. `captured-verbatim`
means the markup is kept as bytes and replayed — right for a historical record
such as a tracked-change snapshot, wrong for anything a user edits. `dropped`
names a reason class from `DROP_REASONS`, each of which states the mechanism and
the kind of fix.

## Where a capture lives

The sink records a capture's position as `index`, a count of the modelled
siblings that preceded it, and `serializeWithPreservedChildren` puts it back
between the same two. That is the right shape for a container whose model is a
list of something else — `Comment.content` is a paragraph list, so a table in a
comment body has nowhere to be except beside an index.

A **block** container is different, and its captures are not in the sink at
all. `BlockContent` gained a `preservedBlock` member, so the capture is a block
in its own right: it sits between the same two siblings in the model, in the
ProseMirror document and in the saved part, and no index has to be kept honest
as the blocks around it are inserted, split, merged or deleted. The editor leg
falls out of that — `preservedBlock` is a zero-width, non-selectable atom whose
position ProseMirror's own mapping maintains — and it is why the block
containers carry no `lost-in-the-editor-projection` losses. The run level is
built the same way: `RunContent.preservedXml` is a member of the run's content
union, not a sink beside it.

The inline level between them is `ParagraphContent`'s `preservedInline`, and
there position is not merely convenient but load-bearing. A paragraph, a
run-level tracked-change wrapper, a bidirectional wrapper and an inline content
control share one walk, and `w:permStart`, `w:proofErr`, `w:customXml` and the
eight custom-XML revision ranges are declared in all of them. Markup lifted out
of a `w:ins` and written beside it survives the save and still breaks the
document: accepting the insertion leaves the markup behind and rejecting it
keeps markup belonging to a change nobody kept. So the capture is a member of
the wrapper's own content union, it rides the insertion mark through the
editor as the same opaque atom the run level uses, and the atom records which
level it came from — `w:ruby` goes back inside a `w:r` and `w:permStart` may
not, because the schema admits no such child of a run.

A link and a simple field are the same level again. `CT_Hyperlink` and
`CT_SimpleField` are both `EG_PContent`, so either may hold a permission
range, a proofing error or a custom-XML revision range between its runs, and
`Hyperlink["children"]` and `SimpleField["content"]` carry `PreservedInline`
for the same reason `ParagraphContent` does. Two things about the link are
worth writing down, because both are easy to get wrong:

- **One map, two callers.** `parseHyperlink` reads a link, and the paragraph
  parser's revision-segmenting walk reads one that holds `w:ins` or `w:del` —
  OOXML nests the revision inside the link and the model nests the link inside
  the revision, so the second cannot simply call the first. The handler map is
  exported and the segmenting caller overrides exactly the four
  `CT_RunTrackChange` names, so the other twenty-nine decisions are made once.
- **The walk is flat and the segmenting happens after it.** A capture the
  dispatcher's sink holds carries an index, and an index counted against
  whichever segment happened to be open when the child was read would place
  the markup in the wrong link. The walk records the hoisted revisions in
  source order as items of the same list, the sink's captures are placed into
  that one list, and only then is it cut into links and revisions.

So the sink's `index` is for a container that models one kind of child, and a
union member is for a container that models a sequence. Prefer the union member
when there is one: an index that has to be maintained is a mirror, and a mirror
drifts.

Either way the pair is `captured-verbatim`, not `modelled`. The law answers the
question by execution — it clears every capture and asks what still survives —
so `CAPTURE_MEMBER_TYPES` and `CAPTURE_SINK_KEYS` in `laws.ts` have to name the
sink's shapes alongside the `rawSomethingXml` fields. A capture that lives in
the model's own union is still bytes, and a contract that called it `modelled`
would promise an editor a thing it cannot edit.

### A table: the same sink, one level up

`CT_Tbl` declares the same markers beside its rows that `CT_Row` declares
beside its cells, and none of them is a row, so `Table.preserved` is the sink
again with `index` counting rows. Three things are worth writing down, because
each of them is a decision and not a consequence:

- **The index counts rows only.** `w:tblPr` and `w:tblGrid` are
  `OWNED_ELSEWHERE`: they precede every row in the content model,
  `serializeTable` writes them from the model ahead of the sink, and the grid
  travels as a capture of its own on the table's formatting. An index that
  counted them would push every capture one place to the right.
- **Three types, one map.** `w:sdt` is unwrapped and its rows spliced into the
  table, so the recursion walks `CT_SdtContentRow` with the same handler map.
  `w:customXml` is not unwrapped: it is captured whole, exactly as the row
  captures a `w:customXml` cell wrapper, which keeps `CT_CustomXmlRow`'s
  <!--count:container=customXml|CT_CustomXmlRow-->31<!--/count--> pairs at the
  price of the wrapper's content being opaque. `CT_CustomXmlRow`
  is a member of the generated set anyway, so its one extra name —
  `w:customXmlPr` — carries a decision rather than falling to a default.
- **A `w:tbl` under a `w:tbl` is captured whole.** The Transitional content
  model does not declare it, so the sink's default already keeps it; that is
  the branch a hand-written `default` gets wrong, and flattening it would move
  its rows into a table nobody wrote.

The editor leg stops at the save law for the reason the row section gives, one
level up: the table node's children are rows, and a zero-width atom between two
of them is not a row. So `tbl|CT_Tbl`'s
<!--count:reason=editorProjection&container=tbl|CT_Tbl&kind=child-->26<!--/count-->
child pairs, and the ones lost with the two row wrappers, are
`dropped (editorProjection)`.

### A row: the sink, and where it stops

`CT_Row` declares a permission range, a proofing error, the row-level comment
and move ranges and the eight custom-XML revision ranges beside its cells.
None of them is a cell, and a row models one kind of child, so this is the
sink case rather than the union case: `TableRow.preserved` holds the capture
with `index` counting the cells that preceded it, and the serializer puts it
back between the same two. `w:trPr` and `w:tblPrEx` are `OWNED_ELSEWHERE` —
the row's property parsers read them off the element, and capturing them as
well would write each twice.

**The editor leg stops here, and the reason is structural.** The block level
carries its captures as a zero-width `preservedBlock` node, so ProseMirror's
own mapping keeps the position honest. The table schema has no row-level node
to do that with: a row's children are cells, and a zero-width atom between two
of them is not a cell. Giving the row one means either a cell-shaped node that
renders nothing — which every command that walks a row would have to learn to
skip — or an attribute on the row node, which is an index and drifts the
moment a column is inserted or deleted.

So a row's captures survive a save and are lost by the editor projection, and
the contract records exactly that: its
<!--count:reason=editorProjection&container=tr|CT_Row&kind=child-->24<!--/count-->
child pairs moved from `dropped (neverParsed)` to `dropped (editorProjection)`.
That is not a lateral
move. `neverParsed` says folio never read the markup and a document that is
merely opened and saved loses it; `editorProjection` says the markup is in the
model and in the saved part, and only a round trip through the editor drops
it. The fix for what remains is one decision about the table schema, not a
parser.

### A transparent inline wrapper: the mark, not the index

`w:bdo`, `w:dir`, `w:smartTag` and the run-level `w:customXml` are all the same
shape. Each is transparent: its children are ordinary inline content and the
wrapper adds a layout control, or a name, a URI and a properties bag. None of
them constrains what it may hold, so dropping one changes what the reader sees
or what the markup states rather than what the text is.

folio used to lose each of them a different way. It spliced a `w:smartTag`'s
children into the paragraph and kept no wrapper, which cost the tag all
<!--count:container=smartTag|CT_SmartTagRun-->37<!--/count--> of its pairs,
the two attributes that identify it among them. It captured a run-level
`w:customXml` whole,
which kept its <!--count:container=customXml|CT_CustomXmlRun-->37<!--/count-->
pairs at the price of every run inside
it being opaque bytes. They are now `InlineWrapper` members beside `bidi`,
discriminated on `kind`, and the content of all four is read by one run-level
walk.

The carrier in the editor is a **mark**, not an index range. The two are worth
telling apart, because the choice is the difference between an editable
wrapper and an opaque one:

- **Why a mark.** A wrapper is not _between_ two children, it is _around_
  several, which is what a union member cannot express and what capturing the
  wrapper whole answers by making every child a child of the capture. A mark
  says the same thing without the capture: marks split and merge with the text
  they are on, so ProseMirror maintains the range and no index has to be kept
  honest as the content around it is edited.
- **One mark, a stack of layers.** `inlineWrapper` is non-exclusive and
  `inclusive: false`, and its `stack` attr lists the wrappers a leaf sits
  inside, outermost first. The multiplicity has to live in the attr because
  ProseMirror's mark set is unordered across types, so two marks could not say
  which wrapper is inside which, and a `w:smartTag` inside a `w:dir` is not the
  same document as a `w:dir` inside a `w:smartTag`.
- **A layer is the model member minus its content.** `WrapperLayer<W>` is
  `Omit<W, "type" | "content">` distributed over the model's union, so a kind
  exists in both places or in neither and the fields a kind carries are stated
  once. A smart tag cannot gain a bidirectional `control` that way, and a kind
  added to the model is a compile error at every site that decides per kind:
  the layer factory, the attrs validator, the DOM spelling, the painter's
  direction and the serializer.
- **The properties are part of the layer.** `w:smartTagPr` / `w:customXmlPr`
  ride the layer as `propertiesXml`, so they are part of the stack key the save
  leg groups by: two adjacent smart tags that differ only in their properties
  are two wrappers, and folding them into one would attribute the second tag's
  text to the first tag's properties. They are bytes, which the contract
  records as `captured-verbatim` rather than `modelled`, and the serializer
  replays them only when they are structurally the element they claim to be —
  a mark is reachable from a paste, and a string that closed the wrapper early
  would splice sibling markup into the part.
- **Save leg.** `fromProseDoc` cuts the paragraph's inline sequence into
  maximal groups of equal stack before it builds runs, and closes the wrappers
  around each group; equal is the layer factory's canonical key, so one wrapper
  never becomes two. Nothing that spans nodes may span the cut, which is what
  makes the cut the right place to finish the open run, hyperlink and revision.
  The order a rebuilt span is written in is fixed — `w:ins` | `w:del` |
  `w:moveFrom` | `w:moveTo`, then the wrapper layers outermost first, then
  `w:hyperlink`, then `w:r` — and the revision is outermost because folio
  already writes one outside the hyperlink it spans, because the parse leg is
  revision-owned, because accepting or rejecting a revision is a range
  operation over its own content, and because a wrapper outside two revisions
  would mint one revision id per wrapper. A group with nothing left in it
  writes no wrapper, so a wrapper whose content was deleted or rejected is
  gone rather than left standing empty.
- **What the canonical order costs.** `w:bdo > w:ins` and `w:ins > w:bdo`
  reach the editor as the same marks on the same leaf, so the save leg cannot
  tell them apart and writes both the canonical way round. A paragraph nobody
  edited keeps its authored order because selective save replays its bytes;
  one rebuilt from the editor — an edit, or a full repack — comes back
  revision-outermost. Recovering the authored order would mean recording it on
  the mark, which is a carrier change and not this one. A tracked edit splits
  the span into a kept part, a deletion and an insertion, and each takes a
  wrapper of its own for the same reason; the wrapper is transparent, so the
  three say what the one said.
- **What unwrapping the custom-XML wrapper cost.** Its children were bytes
  before, and bytes survive anything; now they are modelled, so each of them
  answers the same question the same children of `w:p`, `w:bdo`, `w:dir` and
  `w:smartTag` answer. A revision with a run in it survives; a range marker
  spanning nothing does not, which is a fixture class that
  `lost-in-the-editor-projection` describes below rather than a defect of the
  wrapper.
- **The block level is not this.** A `w:customXml` around two paragraphs is not
  a mark, and needs the index range after all. So the editor leg ships for the
  inline wrappers and the block ones stop at the save law: `w:customXml` as a
  block, a row and a cell wrapper is still captured whole, and the contract
  records that.

### An inline content control: the revision goes around it

An inline `w:sdt` is not a transparent wrapper. It binds its content to
something, so the editor keeps a node for it — `content: "inline*"`, attrs for
the properties, and the `w:sdtPr` / `w:sdtEndPr` bytes beside them. Turning it
into an atom is not on the table: y-prosemirror deletes an element it cannot
rebuild, so the node is what every Yjs document already holds and changing it
is a migration rather than a projection change.

That node is not an inline atom, so it is not a run carrier, and the revision
mark that says its text was inserted has nowhere on the node to sit.
`w:ins > w:sdt` therefore used to reach the editor carrying nothing, and the
save leg wrote the control back beside the revision that had held it: the text
a reviewer inserted was no longer inserted, and accepting the change kept it
exactly as rejecting it did.

The revision now rides the leaves the control holds, which is where
`toProseDoc` already puts it for everything else inside a revision, and the
save leg reads it back off them:

- **A revision that covers all of the content is hoisted around the control.**
  `createInlineSdtFromNode` writes `w:ins > w:sdt`, the canonical order this
  section's save leg already uses for a wrapper and a hyperlink. Outermost is
  also the only form in which accepting or rejecting is an operation over the
  whole control, and the resolution commands agree: rejecting an inserted
  control removes the control rather than leaving an empty one standing where
  its content was. The editor-command path and the headless resolver each
  decide that in their own code, and a test binds them to the same answer.
- **A revision that covers part of the content stays inside.** There is no
  outermost form of "these two runs of five were inserted", so the control
  keeps the revision where the editor holds it, per child.
- **What the canonical order costs, again.** `w:ins > w:sdt` and
  `w:sdt > w:ins` reach the editor as the same marks on the same leaves, so a
  rebuilt span comes back revision-outermost, exactly as `w:bdo > w:ins` does.
  A paragraph nobody edited keeps its authored order, because selective save
  replays its bytes.
- **A wrapper around the whole control keeps its place.** The `inlineWrapper`
  mark sits on the control's node rather than on its leaves, so `w:bdo > w:sdt`
  comes back as it was authored. When a revision and a wrapper both enclosed
  the control, the group is nested inside the revision and the canonical order
  decides the rest: `w:ins > w:bdo > w:sdt`.

### What `lost-in-the-editor-projection` is and is not

<!--count:reason=editorProjection-->247<!--/count--> pairs carry this mechanism,
and reading them as one defect gets the fix wrong. The law compares the
fixture's markup against the part the editor round trip writes, and it asks
only whether the markup is _somewhere_ in that part. Four classes come out of
that, and each is fixed in a different place — or is not a defect at all.

**The table and block sinks: <!--count:reason=editorProjection&container=tbl|CT_Tbl,tr|CT_Row,customXml|CT_CustomXmlRow,customXml|CT_CustomXmlCell,sdtContent|CT_SdtContentRow,sdtContent|CT_SdtContentCell-->160<!--/count-->.**
`w:tbl`, `w:tr` and the block, row and cell wrappers around them hold their
unmodelled children as a capture with an index, and the editor has no node to
carry one: a zero-width atom between two cells is not a cell. The row and table
sections above give the reason in full. This is the largest class and the one
with a single fix — a decision about the table schema — behind it.

**A comment or move range spanning nothing: <!--count:reason=editorProjection&subject=commentRangeStart,commentRangeEnd,moveFromRangeStart,moveFromRangeEnd,moveToRangeStart,moveToRangeEnd&container!=tbl|CT_Tbl,tr|CT_Row,customXml|CT_CustomXmlRow,customXml|CT_CustomXmlCell,sdtContent|CT_SdtContentRow,sdtContent|CT_SdtContentCell-->60<!--/count-->.**
These are the fixture and not folio, and they are deliberate. `PARTNER_MARKERS`
writes a marker's partner _beside_ the subject, because the pair has to be
balanced for the package to measure anything at all; so the range it forms
spans no content. A bookmark range is not in this class and does not need to
be: the editor has a node for a bookmark boundary, so an empty one survives. The editor spells a comment as a mark over inline content and
a move as the kind on a revision mark, and a range over nothing has no carrier
to ride, so the projection drops it. Seeding content between the two markers
instead would change what every marker pair measures, and what a range over
content does is asserted where it belongs, in the conversion and comment tests.
Until that is worth doing, these pairs are `dropped (editorProjection)` with
this paragraph as their reason: the contract's reasons are classes rather than
prose per pair, so a per-pair sentence has nowhere to live in the JSON.

**A revision nested directly inside a different revision: <!--count:reason=editorProjection&container=ins|CT_RunTrackChange,del|CT_RunTrackChange,moveFrom|CT_RunTrackChange,moveTo|CT_RunTrackChange&subject=ins,del,moveFrom,moveTo-->12<!--/count-->.**
The fixture again, and for a reason worth naming. `SEED_CHILDREN` gives a
revision a run, so `w:p > w:ins` measures an insertion with something in it;
but the builder's recursion guard stops at a type already on the path, so the
inner wrapper of `w:ins > w:del` is still written empty. `TrackedRunContent`
admits a nested `TrackedRunChange`, and a nested revision with a run in it
survives. Lifting the guard would change the shape of every fixture that
reaches a depth or recursion limit, which is a wider measurement change than
the seeding was.

**The honest remainder: <!--count:reason=editorProjection&container=hyperlink|CT_Hyperlink,bookmarkStart|CT_Bookmark,bookmarkEnd|CT_MarkupRange,r|CT_R-->15<!--/count-->.**
These are folio.

- <!--count:reason=editorProjection&container=hyperlink|CT_Hyperlink&kind=attribute-->3<!--/count-->
  — `w:hyperlink`'s `w:docLocation`, `w:history` and `w:tgtFrame`. The editor's
  link mark carries `href`, `tooltip` and `rId` and nothing else.
- <!--count:reason=editorProjection&container=bookmarkStart|CT_Bookmark,bookmarkEnd|CT_MarkupRange&kind=attribute-->5<!--/count-->
  — `w:bookmarkStart`'s `w:colFirst`, `w:colLast` and `w:displacedByCustomXml`,
  and `w:bookmarkEnd`'s `w:displacedByCustomXml` and `w:id`. The editor's
  bookmark boundary normalises the pair's position and keeps neither the
  table-column scope nor the displacement.
- <!--count:reason=editorProjection&container=r|CT_R&kind=attribute-->3<!--/count-->
  — `w:r`'s `w:rsid*`. The attribute-remainder section below says why a run is
  the one owner without a record on the other side.
- <!--count:reason=editorProjection&container=r|CT_R&kind=child-->4<!--/count-->
  — `w:softHyphen` and `w:noBreakHyphen`, which the editor carries as U+00AD
  and U+2011 inside the text so the save writes the character rather than the
  element (`present-with-a-different-value` is the truer mechanism; the law
  does not reach it because it looks for the element), and `w:t` and
  `w:instrText`, where the subject _is_ the seed: the fixture writes an empty
  one, and a run holding nothing but an empty text element is an empty run.

**What the law cannot see at all.** The census asks whether the markup is
somewhere in the saved part, so markup lifted _out_ of a wrapper and written
beside it reads as surviving:

```xml
<w:ins …><w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo></w:ins>
<!-- became -->
<w:ins …/><w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo>
```

`x` was no longer inserted, and rejecting the revision kept it. Only a
position-sensitive test sees that, which is why
`trackedWrapperChildSurvival.test.ts` asserts about what is _inside_ the
wrapper rather than what is in the paragraph, and why
`inlineSdtRevisionHoist.test.ts` does the same for the control. Both wrappers
are now admitted by the single total map in `inlineWrapperContent.ts`, and the
revision reaches the editor on the leaves either of them holds. A marker the
same lifting moves is harmless: document order is unchanged and the wrapper
splits into two with the same attributes, which the revision-id pass re-mints.

### The attribute remainder

The child sink is about children. An element's _attributes_ had no branch at
all: a parser read the ones it models off the element and the serializer
rebuilt the start tag from the model, so everything else went. The census
charged
<!--count:kind=attribute&subject=rsidDel,rsidP,rsidR,rsidRDefault,rsidRPr,rsidSect,rsidTr-->20<!--/count-->
`@rsid*` pairs to `w:p`
(<!--count:kind=attribute&container=p|CT_P&subject=rsidDel,rsidP,rsidR,rsidRDefault,rsidRPr,rsidSect,rsidTr-->5<!--/count-->),
`w:r`
(<!--count:kind=attribute&container=r|CT_R&subject=rsidDel,rsidP,rsidR,rsidRDefault,rsidRPr,rsidSect,rsidTr-->3<!--/count-->),
`w:tr`
(<!--count:kind=attribute&container=tr|CT_Row&subject=rsidDel,rsidP,rsidR,rsidRDefault,rsidRPr,rsidSect,rsidTr-->4<!--/count-->)
and `w:sectPr`
(<!--count:kind=attribute&container=sectPr|CT_SectPr&subject=rsidDel,rsidP,rsidR,rsidRDefault,rsidRPr,rsidSect,rsidTr-->4<!--/count-->
each on `CT_SectPr` and `CT_SectPrBase`) as `never-parsed` — every `w:rsid*`
attribute the schema graph declares. Word writes a revision-session id on
nearly every one of those elements, so opening a document and saving it
rewrote the whole revision history.

The remainder is `attributeRemainder.ts`, and four things about it are
decisions rather than consequences:

- **The carrier is the element's own record**, `Paragraph.preservedAttributes`
  and its three siblings, not the child sink: an attribute has no position
  among children to keep. `CAPTURE_SINK_KEYS` in `laws.ts` names the field, so
  the survival law clears it and the contract records `captured-verbatim`
  rather than `modelled` — these are bytes, and an editor cannot edit them.
- **The name is resolved, never the spelling.** folio reads an attribute by
  namespace URI plus local name, with a local-name fallback across prefixes, so
  an element that binds a second prefix to the WordprocessingML URI and writes
  `altw:paraId` is read exactly as `w14:paraId` is. A remainder that matched
  `"w14:paraId"` textually would keep a second copy of an attribute the parser
  had already read, and the save would write the value twice under two
  spellings. So the decision is made on the resolved local name and what is
  written back is the canonical prefix for the resolved URI.
- **Namespace declarations are not in the remainder**, and neither is an
  attribute whose namespace `partNamespaces.ts` cannot spell. That module
  derives a rebuilt part's bindings from the prefixes the part uses and fails
  the save on one it cannot bind; replaying a source element's own `xmlns:*`
  fights the root's, and keeping an attribute under a prefix nothing binds
  would turn a preserved attribute into an unopenable package.
- **The writer, not the predicate, prevents a double.** The modelled attributes
  are handed to `serializePreservedAttributes` as the fragments it is about to
  emit, and a remainder entry that would spell one of them again is dropped.
  The per-owner modelled set mirrors what each parser reads, and a mirror
  drifts; this is what keeps the drift from reaching the part as a duplicate
  attribute.

**The editor leg is the boundary, and the rule is one sentence: the remainder
follows the record, and a record the editor creates has none.** A paragraph, a
row and a section's properties each have a record on the other side —
`ParagraphAttrs._preservedAttributes`, `TableRowAttrs._preservedAttributes`,
and the whole `SectionProperties` object the paragraph attrs already carry —
so an authored element's remainder comes back unchanged. ProseMirror copies a
node's attrs when a command splits it, so both halves of a split paragraph hold
the _same array_, and `keepOneAttributeRemainderPerRecord` gives it to the
first in document order and to no other. Reference identity is what tells the
cases apart, and it is exact: two elements that each parsed their own
attributes hold different arrays however equal their contents, and only a copy
the editor made shares one.

A run is the exception. The editor has no run record: a run is text plus
marks, and a run-level attribute would have to ride a non-exclusive inline
mark. That carrier is the `inlineWrapper` mark, but it carries wrappers rather
than a run's own attributes, and
giving a run its remainder is a separate record with its own grouping rules.
So `r|CT_R`'s
<!--count:reason=editorProjection&container=r|CT_R&kind=attribute-->3<!--/count-->
pairs stay at `editorProjection` — the model holds them and a save writes them.

### Giving `styles.xml` and its neighbours a rebuild law

`w:latentStyles` and `w:lsdException` are not in the census space at all, and
neither is most of `w:style`. `schemaSpace.ts` walks from the roots of the
parts folio rebuilds, and the synthesised fixture is a `w:document`: nothing
reaches `styles.xml`, `numbering.xml`, `settings.xml`, `fontTable.xml` or
`webSettings.xml`. A concurrent branch extends the census to root fixtures at
those five parts and finds that a repack **copies** them byte for byte, so
every pair under them passes without exercising anything. That is not survival,
it is absence of measurement, and the two have to be told apart in the report:
the right word for those pairs is **unmeasured**, not `modelled`.

Making them measurable needs the same forcing the body already has, one level
up. L2 works because `forcedSavePart` strips the verbatim captures the replay
would hand back, so the _element_ serializers run. For these parts the replay
is not a capture inside the model, it is the part itself: `rezip.ts` carries
the original entry across unless something asked for it to be rewritten. So
the law has to force the **part** serializer.

- The part serializers that exist today are `stylesSerializer.ts`,
  `numberingSerializer.ts`, `settingsSerializer.ts`, `fontTableSerializer.ts`
  and `webSettingsSerializer.ts`. Each already takes a parsed model and returns
  a part; none of them is on the save path for an untouched document.
- The census needs a fixture builder rooted at each part (its own content-type
  override and relationship, which the body builder does not synthesise), and a
  `forcedSavePart` variant that calls the part serializer directly instead of
  repacking. The four laws then read the same way they do for the body.
- Expect the result to be large. `w:latentStyles` carries up to 375
  `w:lsdException` children and folio models none of them; `w:style` has a
  wide `w:pPr`/`w:rPr` surface that the style model flattens.
- Latent-style capture then goes where the paragraph's did: on the model
  record for `w:styles`, as a sink of the container's unmodelled children with
  their position among the modelled ones. `w:latentStyles` is a single child of
  `w:styles` with a fixed place in the content model, so one capture holds the
  whole element including its exceptions; splitting it per `w:lsdException`
  would buy nothing, because folio has no model for a single exception either.

### Why totality is a check and not a type

`specifications/reserved-values` proves its totality with `as const satisfies
Record<keyof T, …>`, and that is the better mechanism when the key set is a
model's own fields. Here the key set is about three thousand schema pairs. A
generated union that wide, with a `satisfies` over it, would cost more type
instantiations than every published package put together, and the budget may
not be raised to pay for it. So the keys stay strings, the project measures
what `scripts/typecheck-budget.json` records for it, and totality is enforced
by `bun run check:container-contract`, which costs a census run that has to
happen anyway.

The compiler still owns what a decision may _say_. `ContractEntry` is a
discriminated union, so a `dropped` entry with no reason, or one naming a class
`DROP_REASONS` does not define, fails `bun run typecheck`. And
`REASON_FOR_MECHANISM` in `scripts/container-contract.ts` is
`as const satisfies Record<LossMechanism, DropReason>`: a new way to lose
something cannot be recorded until somebody has given it a name and a fix.

## Adding an element

You do not add it here. Add the parser or serializer support, then run:

```sh
bun run check:container-contract      # names every pair whose behaviour moved
bun run check:container-survival      # the same, per mechanism, against the baseline
```

Each tells you exactly which pairs changed and in which direction. When the
change is the one you meant, record it:

```sh
bun run container-contract:write      # rewrites specifications/container-contract/contract.json
bun run container-survival:baseline   # rewrites specifications/container-contract/survival-baseline.json
bun run container-contract:doc-counts # rewrites the derived numbers in this file
```

The first two are full sweeps and refuse to write from a scoped run, so a
baseline can never be narrowed by accident. During development, scope the
_check_ instead:

```sh
bun scripts/container-contract.ts check --only tblGridChange
bun scripts/container-survival-census.ts run --only numberingChange
bun scripts/container-survival-census.ts fixture --only numberingChange   # see the package it builds
```

## How the baselines shrink

`survival-baseline.json` records, per pair, the mechanism that loses it, and
separately the values a surviving slot still loses. The gate fails when a pair
starts being lost, when a recorded loss changes mechanism — a different defect
wearing the old one's name — and when a recorded loss stops happening without
the baseline being rewritten. The last case is what makes a fix lock its own win
in: the list can only go down.

`contract.json` ratchets the same way from the other side. A pair moving from
`dropped` to `modelled` is a failure until it is recorded, and a pair moving the
other way is a failure full stop.

## Why the numbers in this file are generated

Every count above sits between a marker naming the query that produces it:

```md
<!--count:reason=editorProjection-->247<!--/count--> pairs carry this mechanism
```

`bun run check:container-contract` recomputes each one from `contract.json` and
fails when the file and the contract disagree. The reason is this document's
own history: it claimed 118, 128 and 108 pairs against a contract that held
289. A number typed beside a generated file is a mirror, and a mirror drifts —
the same argument the capture sections make about an index.

A filter is `field=value` or `field!=value`, joined by `&`, with a
comma-separated value list meaning "any of". The fields are the contract's own:
`disposition`, `reason`, `kind`, `container` and `subject`, each matched against
the pair key with its namespaces removed. A field the script does not define, a
value no pair carries, and a query that selects nothing are all errors rather
than zeroes: a marker matching nothing is a stale key, and reading it as zero is
how the drift started.

## Where this came from

The public corpus census (423 failure signatures over 3567 packages) found the
instances: `w:moveFromRangeStart` reaching disk without the `w:author` its
schema type requires, `w:tblGridChange` disappearing whenever a column was
resized, `w:numberingChange` disappearing always. The instances were fixed. The
class — a parser that hands an unrecognised child to the floor — needed a census
over the schema rather than over examples, because the corpus can only show what
some document happened to contain.

See also `docs/reserved-values.md`, which holds the same shape for a different
question: there the registry is total over the model and the coverage check is
total over the schema; here the contract is total over the schema and the law is
what verifies it.
