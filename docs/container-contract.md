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
leave a stale confession behind and the list would only ever grow.

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

The graph serialises a container's particles sorted by id, which is
lexicographic, so declaration order has to be read off each particle's `order`.
Three readers depend on the result — the fixture builder, which puts the
subject at the ordinal its particle declares; `corpus-schema-validator.ts`,
which scores a document's children against the same sequence; and
`generate-container-children.ts`, whose `sequence` rows become the order a
property-set serializer writes. They share one derivation
(`orderedParticlesByOwner` in `scripts/lib/ooxml-schema-graph.ts`), and
`scripts/container-children-order.test.ts` asks the validator whether the
generated order is still the order it scores against, because three copies of
one sort are three chances to disagree.

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
  no paragraph — 212 pairs were charged to the wrong container for exactly that
  reason.
- `SEED_CHILDREN` in `fixture.ts` gives a container the content it needs to
  survive at all — a row in a table, a paragraph in a cell, a numbering
  reference in a `w:numPr`. A container folio prunes for being empty would
  report every pair inside it as lost.

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

- **The index counts rows only.** `w:tblPr` and `w:tblGrid` are owned
  elsewhere: they precede every row in the content model,
  `serializeTable` writes them from the model ahead of the sink, and the grid
  travels as a capture of its own on the table's formatting. An index that
  counted them would push every capture one place to the right.
- **Three types, one map.** `w:sdt` is unwrapped and its rows spliced into the
  table, so the recursion walks `CT_SdtContentRow` with the same handler map.
  `w:customXml` is not unwrapped: it is captured whole, exactly as the row
  captures a `w:customXml` cell wrapper, which keeps `CT_CustomXmlRow`'s 29
  pairs at the price of the wrapper's content being opaque. `CT_CustomXmlRow`
  is a member of the generated set anyway, so its one extra name —
  `w:customXmlPr` — carries a decision rather than falling to a default.
- **A `w:tbl` under a `w:tbl` is captured whole.** The Transitional content
  model does not declare it, so the sink's default already keeps it; that is
  the branch a hand-written `default` gets wrong, and flattening it would move
  its rows into a table nobody wrote.

The editor leg stops at the save law for the reason the row section gives, one
level up: the table node's children are rows, and a zero-width atom between two
of them is not a row. So `tbl|CT_Tbl`'s 26 child pairs, and the 55 that were
lost with the two row wrappers, move to `dropped (editorProjection)`.

### A row: the sink, and where it stops

`CT_Row` declares a permission range, a proofing error, the row-level comment
and move ranges and the eight custom-XML revision ranges beside its cells.
None of them is a cell, and a row models one kind of child, so this is the
sink case rather than the union case: `TableRow.preserved` holds the capture
with `index` counting the cells that preceded it, and the serializer puts it
back between the same two. `w:trPr` and `w:tblPrEx` are owned elsewhere —
the row's property parsers read them off the element ahead of the cells, and
capturing them as well would write each twice. For `w:tblPrEx` that claim was
false until the property-set section below gave it an owner, which is why an
owner claim now has to name the reader it claims.

**The editor leg stops here, and the reason is structural.** The block level
carries its captures as a zero-width `preservedBlock` node, so ProseMirror's
own mapping keeps the position honest. The table schema has no row-level node
to do that with: a row's children are cells, and a zero-width atom between two
of them is not a cell. Giving the row one means either a cell-shaped node that
renders nothing — which every command that walks a row would have to learn to
skip — or an attribute on the row node, which is an index and drifts the
moment a column is inserted or deleted.

So a row's captures survive a save and are lost by the editor projection, and
the contract records exactly that: the 24 child pairs move from
`dropped (neverParsed)` to `dropped (editorProjection)`. That is not a lateral
move. `neverParsed` says folio never read the markup and a document that is
merely opened and saved loses it; `editorProjection` says the markup is in the
model and in the saved part, and only a round trip through the editor drops
it. The fix for what remains is one decision about the table schema, not a
parser.

Both are transparent: their children are ordinary inline or block content and
the wrapper adds a name, a URI and some properties. folio splices a
`w:smartTag`'s children into the paragraph and keeps no wrapper, which costs
`w:smartTag` its own 29 pairs and the two attributes that identify it.
`w:customXml` is now captured whole instead, which keeps its 35 pairs at the
price of its content being opaque in the editor — the right trade only because
that content was previously dropped outright.

Neither is the end state. The end state is a **`preservedWrapper`**: a range
over the container's child indices, recorded beside the children rather than
instead of them.

- The record is `{ xml: string; from: number; to: number }` where `xml` is the
  wrapper's start tag plus its `w:customXmlPr` / `w:smartTagPr` and its
  attributes, and the two indices bracket the modelled children it held. On
  save the serializer re-opens the wrapper before the child at `from` and
  closes it after the child at `to`, so the children stay modelled and
  editable and the wrapper comes back in the authored position.
- Nesting falls out of ranges: two wrappers over overlapping-but-nested spans
  re-open in index order, outermost first, which is the order they were read
  in. Overlapping-but-not-nested ranges cannot occur, because the source was
  a tree.
- This is the one place the contract's "prefer a union member, an index
  drifts" rule does not apply, and it has to be said why: the wrapper is not
  _between_ two children, it is _around_ several, and a union member cannot
  express that without making every child a child of the capture — which is
  what capturing the wrapper whole already does, and is what costs the editor
  the content.
- The index does drift, and that is the honest cost. An edit that inserts a
  paragraph inside the range grows the range in a way the author did not
  write, and an edit that deletes every child in it leaves an empty wrapper.
  Both are recoverable (the range clamps, an empty wrapper is still valid
  markup); neither is losing content, which the alternatives are.
- **Editor leg.** The carrier exists, and it is one mark for every kind of
  inline wrapper rather than one per kind: `inlineWrapper`, non-exclusive and
  `inclusive: false`, whose `stack` attr lists the wrappers a leaf sits
  inside, outermost first. The multiplicity has to live in the attr because
  ProseMirror's mark set is unordered across types, so two marks could not say
  which wrapper is inside which. Marks split and merge with the text they are
  on, so the range is maintained by ProseMirror rather than by an index. Only
  `bidi` is a layer kind today; a smart tag and a custom-XML wrapper are added
  members of the same union with the same attr, and the validator and the
  serializer are total over the kinds so neither can land without a decision.
  It stops being cheap at the block level: a `w:customXml` around two
  paragraphs is not a mark, and needs the index range after all. So the editor
  leg ships for the inline wrappers with the mark, and the block ones stop at
  the save law and the contract says so.

### A property set: where the sink's index stops being a count

`w:tblPr` and `w:sectPr` are not lists of one kind of child, and they are not
sequences of blocks either. Each is a **fixed sequence of optional
singletons**: the schema declares seventeen children for `CT_TblPrBase` and
twenty-two for `CT_SectPr`, every one of them `minOccurs="0" maxOccurs="1"`,
and a consumer refuses a property set whose children are out of that order.
Three things follow, and each is a decision.

- **The sink's index is the schema ordinal, not a count of modelled
  siblings.** A count is a mirror of whichever properties folio models today:
  model one more of them and every capture after it moves one place. Position
  in a sequence is a property of the name, so `sequencePositions` reads it off
  the generated declared-child list, and `serializeSequenceChildren` merges the
  modelled and the captured halves by the same key. An undeclared child — a
  foreign namespace, an `mc:` construct — has no place in the sequence and
  takes the place of the last declared child before it.
- **The order is the generated list, not the order of the serializer's
  statements.** `serializeTableFormatting` used to carry the sequence in a
  comment above thirteen `if` blocks. A comment is a mirror too. The generator
  marks a row `sequence: true`, emits its names in declaration order instead of
  sorted, and refuses two members that order a shared child differently — so
  the set the handler map is total over and the order the serializer writes are
  one list.
- **A handler answers with what it took.** `<w:cols/>` states no column count
  and `<w:jc w:val="end"/>` states a value the reader's enumeration does not
  admit. Neither can be decided by a map keyed on the child's name: the map can
  name the elements folio has never heard of, not the values a reader refuses.
  So `ChildDisposition` lets a handler return `CAPTURE`, meaning "I looked and
  took nothing from this", and the bytes go to the sink. That is what moved the
  35 `never-parsed` pairs, and it is why the fixture generator needed no change
  for them: the census wrote what the schema allows, and folio was losing it.

Two of the 16 section pairs were the fixture rather than folio.
`w:headerReference` and `w:footerReference` were built with an `r:id` no part
in the package answered, and folio removes a dangling reference — which is
right, because that is what makes Word offer to repair the file. The fixture
now carries the header or footer part the reference names, the way it already
carried the footnote a `w:footnoteReference` names.

The editor leg follows the record. `TableAttrs._originalFormatting` carries the
whole `TableFormatting` through ProseMirror and a section's properties travel
whole as well, so the sink and the four newly modelled properties ride them
with no new attr: the pairs move to `captured-verbatim` and `modelled` rather
than stopping at the save law.

#### The third property set, and the decision nobody made

`w:tblPrEx` is the table properties a row overrides, and it had no owner at
all. The row's child walk called it `OWNED_ELSEWHERE` — "another reader owns
this child, and re-emits it" — and no reader did. That is the one way a stated
disposition can still be a silent drop: `CAPTURE` and a handler are checked by
running them, and `OWNED_ELSEWHERE` is a claim about a second place in the
code. Twenty-four pairs went with it, from the row pair down through
`CT_TblPrEx`, `CT_TblPrExBase` and `CT_TblPrExChange`. The section after this
one is about the class rather than the instance.

`CT_TblPrEx` is the middle of `CT_TblPrBase` — the nine properties a row may
restate — and three things follow:

- **One set of handlers, two containers.** The shared children are read by the
  same functions into the same `TableFormatting`, so an exception and the
  property it overrides cannot be read into two different shapes, and
  `TableRow.tablePropertyExceptions` is that record. The generated set is a row
  of its own rather than two more members of `table-properties`: the union
  would make the exceptions' handler map total over `w:tblStyle`,
  `w:tblCaption` and six more names a `w:tblPrEx` cannot hold, and its parser
  would then record decisions for children that never reach it — the reason
  `w:hyperlink` and `w:fldSimple` each have a row too.
- **Declaration order decides the write position.** `CT_Row` is
  `w:tblPrEx, w:trPr, (cells)*`, so the element is written ahead of the row's
  own properties; a serializer that put it after them produces a row a
  validating consumer refuses. The row's sink still counts cells and nothing
  else, exactly as `w:tblPr` and `w:tblGrid` leave the table's counting rows.
- **An empty element is kept.** `w:tblPr` is required on a `w:tbl` and written
  back whatever the model holds, so a reader that takes nothing from it costs
  nothing. `w:tblPrEx` is optional, which makes its presence the value: a
  parser that returned "no exceptions" for `<w:tblPrEx/>` deleted the element.
  The record is present-and-empty rather than absent.

The editor leg follows the record, as the table's does.
`TableRowAttrs._tablePropertyExceptions` carries the set and `tblPrExChange`
the revision, so all 24 pairs come back from the round trip rather than
stopping at the save law the way the row's cell-sink captures do: 16
`modelled` and 8 `captured-verbatim`. Which 8 is worth noting, because it is
not a property of the exceptions at all. `w:tblBorders`, `w:tblCellMar`,
`w:tblLayout` and `w:tblLook` carry `captured-verbatim` on both `CT_TblPrEx`
and `CT_TblPrExBase`, and they carry it on `w:tblPr` too — one set of handlers
gives one answer per child, so the exceptions and the properties they override
are recorded alike.

#### An owner claim names its owner

The instance was one missing parser; the class is a disposition nothing runs.
So `OWNED_ELSEWHERE` stopped being a word. `ownedElsewhere` in
`docx/containerChildren.ts` is now the only way to make one and it takes the
reader as `<module>#<export>`, which turns the claim into data three checks can
read.

- **The child cannot drift from the key.** `ChildHandlers` is keyed per child,
  so an entry filed under `trPr` that names `tblPrEx` does not compile. The
  claim restates the container and the child, and the compiler makes the
  restatement free of doubt rather than a mirror somebody maintains.
- **The reader has to be there.** `scripts/container-ownership.test.ts` imports
  every module a claim names and looks the export up. A reader that was renamed
  or never written fails the test instead of sitting in a comment.
- **The contract has to agree.** `scripts/lib/container-survival/ownership.ts`
  expands each claim over its dispatcher row's members and refuses a pair the
  contract records `dropped (neverParsed)` or `dropped (containerNotKept)`.
  Those are the two reasons an owner contradicts: the first says no parser
  reads the markup, the second says the pair went with a container folio does
  not keep, and in neither case can an owner have written it back. The rest —
  `replayOnly`, `editorProjection` and their neighbours — describe markup a
  reader did take and a later stage lost, which is what the claim says
  happened. `CONTRADICTS_AN_OWNER` is total over `DropReason`, so a new reason
  class is classified rather than defaulting to "not a contradiction".

A claim is registered when the module that makes it loads, so every claim is at
module scope and the check loads the claiming modules by scanning the sources
for the call rather than by keeping a list beside them. The check reads the
committed contract and runs no census, so it costs nothing and runs under
`--only` too: a scoped run must not be a way to land a claim with nothing on
the other end of it. `bun run check:container-contract` runs it.

`KNOWN_OWNED_LOSSES` records a claim that outruns its owner, with the
mechanism, shrink-only the way the survival baseline is — the check also fails
on an entry that has stopped violating. It is empty. The two entries it held
were `tr|CT_Row/trPr` and `tc|CT_Tc/tcPr`, and the section below is what closed
them.

#### The last two property sets, and the element as the carrier

`w:trPr` and `w:tcPr` were the last property sets read by name: ten of the
row's fifteen declared children had an `if` and the rest had nothing, so
`w:cnfStyle`, `w:divId` and `w:tblCellSpacing` went on every save, as did
`w:hMerge`, `w:headers` and the cell's structural revision one level down.
Both sets are dispatcher rows now, with `TableRowFormatting.preserved` and
`TableCellFormatting.preserved` as the sinks and one writer each through
`serializeSequenceChildren`. Four things are worth writing down.

- **A row's properties have no order; a cell's do.** `CT_TrPrBase` is a
  repeated `choice`, so the twelve properties may be written in any order and
  the generated list is *a* valid order rather than the only one — which is
  still better than the order of the serializer's statements, because it is
  derived. `CT_TcPrBase` is a sequence and every type extending it is one, so
  a `w:tcPr` in any other order is markup Word refuses. `CT_TrPr` closes a
  sequence over its base either way, so `w:ins`, `w:del` and `w:trPrChange`
  come last, as `w:tcPrChange` does.
- **The corpus validator scores neither.** `contentModelFor` refuses an order
  wherever the model can reorder itself, and `CT_TrPrBase` is a choice while
  `CT_TcPr` reaches `EG_CellMarkupElements` through its extension chain. A
  clean verdict from it therefore means both "in order" and "there is no order
  to be in", and `container-children-order.test.ts` read the two as one: its
  anti-vacuity guard asked that *some* member of a row refuse a reversal, which
  no member of these two can. It now asks `ordersChildrenOf` first and pins
  every member to a definite verdict, so a validator that stops ordering
  `CT_TblPrBase` fails rather than passes. What it no longer does for these
  two rows is check the order at all, so
  `tableCellPropertySet.property.test.ts` carries that assertion instead:
  every declared child authored backwards, through the serializer, read back
  against the schema's order.
- **An empty element is kept**, the decision `w:tblPrEx` already made. Both
  elements are optional, so presence is the value, and a parser that returned
  "no properties" for `<w:tcPr/>` deleted what a producer wrote. The evidence
  for the other rule is not there: across the 54 packages in the tree — 60
  `w:trPr`, 2941 `w:tcPr`, 64 `w:tblPrEx` — not one of the three is ever
  written empty, so the scan separates neither set from the one already
  keeping its empties. The carrier is the element rather than the properties
  it yielded, which is what closed both `KNOWN_OWNED_LOSSES` entries.
- **A snapshot's structural revision belongs to the snapshot.**
  `CT_TcPrInner` declares `EG_CellMarkupElements`, so the `w:tcPr` inside a
  `w:tcPrChange` may state "before this change the cell stood inserted".
  Capturing it in the walk would have written it twice on the cell's own
  property set, where `parseTableCell` already owns it, so
  `TableCellPropertyChange.previousStructuralChange` carries it and the owner
  claim is true on both members of the row.

The editor leg needed no new attr for the walk itself:
`TableRowAttrs._originalFormatting` and `TableCellAttrs._originalFormatting`
carry the whole record, so the sinks ride them and the pairs are `modelled` and
`captured-verbatim` rather than stopping at the save law. It did need one for
the width. `TableCellAttrs.width` is the width the cell *renders* at, which the
table resolves from its grid when the cell declares no `w:tcW`, and the way back
wrote it into `w:tcPr` unconditionally, so a cell that stated none acquired one.
`TableCellAttrs._authoredWidth` is the record of what the cell states, as
`_resolvedBorders` and `_resolvedMargins` are for the border and the margin, and
the save leg writes `w:tcW` from it alone. A command that moves a cell's width
states one: `mergeTableCellAttrs` derives the record for every command that
patches a cell, so a resize writes exactly the cells it moved.

### What `lost-in-the-editor-projection` is and is not

138 pairs carry this mechanism, and reading them as one defect gets the fix
wrong. The law compares the fixture's markup against the part the editor round
trip writes, and it asks only whether the markup is _somewhere_ in that part.
Two things follow, and they point in opposite directions.

**The census over-reports.** 108 of the 138 are the fixture rather than folio.
A fixture puts the subject in the cheapest container that will hold it, which
for these means an empty one: an empty `<w:ins/>` inside another, a comment
range whose comment the fixture never writes, a move range with nothing moved,
and the `w:author` / `w:date` / `w:id` of a wrapper holding no run. The editor
spells a run-level revision as a _mark on inline content_ and a comment as a
range over it; markup with no content under it has nothing to carry it, and
dropping it is the projection working. `TrackedRunContent` already admits a
nested `TrackedRunChange`, so a non-empty one survives. These are `dropped`
with reason `editorProjection`, and the reason is the fixture's emptiness, not
a missing projection.

The honest remainder is 30:

- **18 + 2** — `w:bdo` and `w:dir` in each of the nine containers that declare
  them, plus their `w:val`. The projection now records both on the
  `inlineWrapper` mark of the leaves the wrapper held, so the direction
  reaches the editor and the painter. They stay on this list until the save
  leg rebuilds the wrapper from the mark instead of from the source
  paragraph: what the census measures is the round trip, and only the read
  half is done.
- **3** — `w:hyperlink`'s `w:docLocation`, `w:history` and `w:tgtFrame`. The
  editor's link mark carries `href`, `tooltip` and `rId` and nothing else.
- **5** — `w:bookmarkStart`'s `w:colFirst`, `w:colLast` and
  `w:displacedByCustomXml`, and `w:bookmarkEnd`'s `w:displacedByCustomXml` and
  `w:id`. The editor's bookmark boundary normalises the pair's position and
  keeps neither the table-column scope nor the displacement.
- **2** — `w:softHyphen` and `w:noBreakHyphen` in a run. These are not lost:
  the editor carries them as U+00AD and U+2011 inside the text, and the save
  writes the character rather than the element. `present-with-a-different-value`
  is the truer mechanism; the law does not reach it because it looks for the
  element.

**The census also under-reports, and that is the more serious half.**
`pushTrackedChangeSegments` lifts out of the wrapper everything
`TrackedRunContent` does not admit, and writes it beside. For a _marker_ —
a comment range, a move range — that is invisible and harmless: document order
is unchanged and the wrapper simply splits into two with the same attributes,
which the revision-id pass then re-mints. For a _content-carrying wrapper_ it
changes the document:

```xml
<w:ins …><w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo></w:ins>
<!-- becomes -->
<w:ins …/><w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo>
```

`x` is no longer inserted. Rejecting the revision now keeps it. `w:dir` and an
inline `w:sdt` do the same thing. The law cannot see it, because the markup is
still in the part; only a position-sensitive test can, which is why
`trackedWrapperChildSurvival.test.ts` asserts about what is _inside_ the
wrapper rather than what is in the paragraph.

The fix is to widen `TrackedRunContent` (and `InlineSdt["content"]`) through
the single total map in `inlineWrapperContent.ts`. Half of what blocked it is
gone: the transparent wrapper has the non-exclusive mark the
`preservedWrapper` section proposes, so a wrapper inside a revision reaches
the editor with both the revision mark and its own stack on the same leaf. An
inline content control is still an `inline*` node rather than an atom, so a
revision mark applied to it lands on its children instead of on the control.

### The attribute remainder

The child sink is about children. An element's _attributes_ had no branch at
all: a parser read the ones it models off the element and the serializer
rebuilt the start tag from the model, so everything else went. The census
charged 20 `@rsid*` pairs to `w:p` (5), `w:r` (3), `w:tr` (4) and `w:sectPr`
(4 each on `CT_SectPr` and `CT_SectPrBase`) as `never-parsed` — every `w:rsid*`
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
mark. That carrier now exists in the shape the `preservedWrapper` section
described, but it carries wrappers rather than a run's own attributes, and
giving a run its remainder is a separate record with its own grouping rules.
So `r|CT_R`'s three pairs stay at `editorProjection` — the model holds them
and a save writes them.

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
not be raised to pay for it. So the keys stay strings, the project measures 168
types and 4 instantiations, and totality is enforced by
`bun run check:container-contract`, which costs a census run that has to happen
anyway.

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
```

Both are full sweeps and refuse to write from a scoped run, so a baseline can
never be narrowed by accident. During development, scope the _check_ instead:

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
