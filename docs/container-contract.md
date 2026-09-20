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
  back, save, and find the subject in the saved part under the chain the fixture
  wrote it at, as many times as the fixture wrote it, with an equal value.
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
| `repeat-truncated`                    | The slot comes back with fewer instances than were written into it.                        |
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

### Where the law looks, and how many it wants

The probe is the law. It used to ask whether the subject's name appeared
anywhere in the saved part, and both halves of that question were wrong.

- **Anywhere.** A pair is (container, child). The `w:pgSz` of the section
  snapshot a `w:sectPrChange` holds is not the live section's `w:pgSz`, and a
  part that lost the first still contains the second. The same goes for every
  wrapper folio unwraps: a `w:smartTag`'s runs are spliced into the paragraph,
  a `w:bdo`'s content reaches the editor without the wrapper, a row-level
  `w:bookmarkStart` is re-anchored inside a cell's paragraph, and a marker
  hoisted out of a `w:ins` is written beside it. In each case the element is
  still in the part, at a place the source did not put it.
- **How many.** `CT_WrapPath` declares `minOccurs="2"` on `wp:lineTo`, so a
  reader that keeps the first vertex and drops the second passes a probe that
  counts nothing.

So the probe walks the saved part tracking the chain of element names above
each start tag, and counts the occurrences that sit **under the fixture's own
container path**: the innermost container is the occurrence's parent, and the
ancestors above it appear in order from the part root. The ancestors are a
subsequence rather than an exact chain, because a save may legitimately wrap
what it writes; the parent and the order are what say the element came back
where it was written. The chain is read from the same container space
`fixture.ts` builds the package from, so where the law looks and where the
fixture wrote are one derivation.

How many it wants is measured, not declared: the same probe counts the
instances the fixture itself placed under that chain, bounded by the slot's
`maxOccurs`. The bound is what keeps a generated fixture honest — a seed and a
subject can land on the same particle, so `w:numPr` gets the `w:ilvl`/`w:numId`
pair that makes it a list plus the `w:numId` under test, and keeping the one
`w:numId` the schema admits is not a loss.

A shortfall is `repeat-truncated`. It is its own mechanism because it is its own
fix — a reader or serializer that handles one instance of a repeated particle
and not the rest — and because it can only be observed where the container came
back, so it never competes with `the-container-itself-is-lost`. No pair carries
it today: the seventeen fixtures that write a two-vertex wrap polygon get both
vertices back, which is the first time anything has said so. It exists so the
first one that does not is named rather than counted as equal.

The container probe asks its question the same way, so a pair whose container is
missing _at that chain_ is `the-container-itself-is-lost` even when an element
of that name survives elsewhere. And `explain` prints the chain it searched with
the counts it found, because a pair reported lost that the printed markup
plainly contains is a pair found somewhere else.

### What asking _where_ moved

Re-measuring the whole space moved 189 pairs, every one of them to an equal or
weaker disposition, and nothing the other way: a probe that asks where can only
refuse what a probe that asks whether accepted. The value sweep and the
unrepresentable counts are unchanged.

| Pairs | Now                                   | What the old probe was finding instead                                                                                                                                                                                                                                                                                |
| ----- | ------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 97    | `the-container-itself-is-lost`        | A wrapper folio unwraps: `w:smartTag` (32), a `w:sdt` around rows or cells (54), and a property snapshot (11: a `w:pPrChange`'s `w:pPr`, a `w:tblPrChange`'s `w:tblPr` and that `w:tblPr`'s nine children). The name was the live element's.                                                                          |
| 42    | `lost-in-the-editor-projection`       | The children of `w:bdo`, `w:dir` and a `w:customXml` row or cell wrapper. The content reaches the editor; the wrapper does not, so its children come back as somebody else's.                                                                                                                                         |
| 44    | `parsed-but-not-serialized`           | Markup the save writes beside its container rather than inside it: the markers `pushTrackedChangeSegments` hoists out of a revision wrapper (34, counting the link that holds one) and a `w:bookmarkStart`/`w:bookmarkEnd` on a body, cell, row or block content control, which is re-anchored into a paragraph (10). |
| 6     | `serialized-only-via-verbatim-replay` | `wp:wrapSquare`/`wp:wrapTopAndBottom`'s own `wp:effectExtent`, written on the `wp:anchor` the way the insets were, and the header and footer bookmarks.                                                                                                                                                               |

The contract moved 196: the same 189 plus seven the carrier probe re-read. A
container nested in one of its own kind — `w:hyperlink` in a `w:hyperlink`,
`w:fldSimple` in a `w:fldSimple`, `w:r` in `w:rt`, `w:rubyBase` and
`w:customXml`, `w:gridCol` in `CT_TblGridBase`, `w:tblGrid` in a
`w:tblGridChange` — was read as `modelled` because a model-only save still wrote
the outer one. What each of them is carried by, and why, is
[below](#a-container-nested-in-one-of-its-own-kind).

Two of these groups are worth a decision rather than a fix. The marker hoist is
deliberate: document order is unchanged and the wrapper splits in two, which is
the behaviour the section below describes as harmless. It is recorded as
`dropped (parsedNotSerialized)` because that is where the markup ends up, and
the alternative is a probe that knows which relocations folio meant — the
leniency this change removed. The bookmark re-anchoring is the same shape and
is not harmless: a bookmark that spanned a row comes back inside one cell's
paragraph.

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
  reason. `w:document` is the same shape one level up: its spine is `w:body`
  and its other child is `w:background`, a page backdrop the schema lets hold a
  `w:drawing`. That chain is two steps and `w:body`/`w:p`/`w:r`/`w:drawing` is
  four, so all 118 DrawingML pairs — `wp:anchor`, `wp:inline`, `wp:docPr`,
  `a:hlinkClick`, `a:hlinkHover`, the five wrap modes — measured how folio
  treats a document backdrop rather than a picture in a run, and every one of
  them read as `the-container-itself-is-lost`. `background|CT_Background` keeps
  its own pairs and still records what folio does with the direct form.
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
"{…}hyperlink|{…}CT_Hyperlink/{…}hyperlink": { "disposition": "captured-verbatim" }
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

### A container nested in one of its own kind

Seven pairs are a container the schema lets hold another of itself, and the
carrier probe reads every one of them by asking what a save with nothing
verbatim left still writes. That is the question worth asking here, because the
failure is not that the markup disappears — it is that the markup is somebody
else's. A parser that splices the inner container's runs into the outer one
writes a part containing every element name the source had, with the inner
container's content attributed to the outer.

The ladder is the same one the rest of this document walks: model the nesting
when the schema and Word both produce it, otherwise capture it into the owner's
own content union at the position it was read at. What separates the two groups
here is measurement rather than taste, so the counts are the public corpus's
(5299 readable packages).

**Modelled.** `w:tblGrid` in a `w:tblGridChange` and its `w:gridCol` children
are the grid a reviewer replaced when resizing a column: 28 packages carry the
element, 104 columns between them. It travelled as `gridChangeXml`, so the
snapshot existed only as markup nothing read and the rebuild could only copy
the string back. `TableFormatting.gridChange` is the parsed shape — the
revision's `@w:id` and one entry per column. `w:w` is optional on a
`w:gridCol`, so a column the snapshot stated no width for is `undefined` and
comes back without one: a snapshot records what stood, and a column with no
measure is not a column of width zero.

Four pairs move to `modelled`: the two from the seven, the `w:tblGridChange`
the live `w:tblGrid` holds, and that change's own `@w:id`. The last two were
already recorded as captured, so the move is a strengthening rather than a
correction.

**Captured, in the owner's union.** The remaining five are kept rather than
modelled, in four groups, each for its own measured reason:

- **`w:hyperlink` in a `w:hyperlink`** — `CT_Hyperlink` is `EG_PContent`, so
  the schema admits it, and the corpus has four occurrences in one package,
  written by a converter rather than by Word. Modelling a link whose own
  children are links means a recursive link mark the editor has no shape for,
  for markup no producer writes.
- **`w:fldSimple` in a `w:fldSimple`** — a nested field is a real thing, and
  Word does not spell it this way: the corpus has zero of them, because a field
  whose result holds a field is written with `w:fldChar` runs. Modelling it
  would widen the editor's `structuredField` — an atom whose content expression
  names text, breaks, images and captures — to hold a field, for a shape
  nothing produces.
- **`w:r` in a `w:rt` and in a `w:rubyBase`** — phonetic guide text, 45 and 46
  runs across 6 packages, and genuine Word output. The whole `w:ruby` is one
  `RunContent.preservedXml` member, so the two pairs are inside a capture that
  keeps its position in the run's content. Modelling it means an editor carrier
  for a reading printed above a base, which does not exist; capturing the
  wrapper whole is what keeps the base's text on the line today.
- **`w:r` in a run-level `w:customXml`** — a transparent wrapper, captured
  whole as a `preservedInline` member. The end state is the `preservedWrapper`
  above, which keeps the children modelled and the wrapper beside them; until
  it exists, capturing the wrapper is what keeps its 35 pairs at the price of
  its content being opaque.

A capture is only the weaker half of the ladder if it keeps the bytes and
nothing else. All of them keep their position, and all of them now keep the
text a reader sees: the two transparent wrappers already did, and the nested
link and the nested field were captured through the child sink, which knows nothing
about a capture beyond its bytes. They are captured through the element
instead, so `preservedRunContent`'s visible-text table answers for them the way
it does for `w:customXml` and `w:smartTag`, and a linked clause inside a link
no longer reaches the editor as an atom showing nothing.

`nestedSameKindContainers.test.ts` is where the class is pinned, and it asserts
about the inner element's *parent* rather than about the part containing a
name, because the flattening failure passes every probe that only counts.

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
- **Editor leg.** The cheap version is the bidi one: `w:bdo`/`w:dir` already
  reach the editor as a mark spanning the inline content they wrap, and a
  smart tag or custom-XML wrapper is the same shape — a non-exclusive mark
  carrying the opaque start-tag markup, applied to every inline node in the
  range. Marks split and merge with the text they are on, so the range is
  maintained by ProseMirror rather than by an index. It stops being cheap at
  the block level: a `w:customXml` around two paragraphs is not a mark, and
  needs the index range after all. So the editor leg should ship for the
  inline wrappers with the mark, and the block ones should stop at the save
  law and say so in the contract.

### What `lost-in-the-editor-projection` is and is not

138 pairs carried this mechanism when the section was written, and reading them
as one defect gets the fix wrong. The law compares the fixture's markup against
the part the editor round trip writes; it used to ask only whether the markup
was _somewhere_ in that part, and now asks whether it is under the container it
was written into. Two things follow, and they point in opposite directions.

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
  them, plus their `w:val`. The editor has no bidirectional mark, so
  `withoutBidiWrappers` keeps the content and loses the direction.
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
inline `w:sdt` do the same thing. The law now sees the hoist for the markers —
they are `parsed-but-not-serialized`, because the probe looks under the wrapper
rather than across the paragraph — and still not for these three, whose markup
comes back inside the wrapper on the save leg and is lost one step later, in the
editor. Only a position-sensitive test reaches that, which is why
`trackedWrapperChildSurvival.test.ts` asserts about what is _inside_ the
wrapper rather than what is in the paragraph.

The fix is to widen `TrackedRunContent` (and `InlineSdt["content"]`) through
the single total map in `inlineWrapperContent.ts`, and it is blocked on one
decision rather than on effort: the editor has no carrier for either wrapper.
A bidirectional wrapper wants the same non-exclusive mark the
`preservedWrapper` section proposes for a smart tag, and an inline content
control is an `inline*` node rather than an atom, so a revision mark applied to
it lands on its children instead of on the control. Widening the model without
those two is a save-leg fix with an editor leg that undoes it on the first open.

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

A run is the exception, and for the reason the bidirectional wrapper is. The
editor has no run record: a run is text plus marks, and a run-level attribute
would have to ride a non-exclusive inline mark, which is the same undesigned
carrier the `preservedWrapper` section is blocked on. So `r|CT_R`'s three
pairs move from `neverParsed` to `editorProjection` — the model holds them and
a save writes them — and they stay there until that mark exists.

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

### Giving a drawing a rebuild law

The same absence of measurement had a second home, one level down. A `w:drawing`
the fixture synthesised carried no `a:blip`, so it named no picture
relationship, so `runParser.ts` classified it preserve-only and the save replayed
its captured bytes whatever the serializers would have written. All 83 pairs
under `wp:inline`, `wp:anchor`, `wp:docPr` and the wrap elements passed on the
strength of a byte copy, and `serializeDrawingContent` — the path every edited
drawing takes — was never run at all.

`fixture.ts` now seeds the `a:graphic` of a drawing with a `pic:pic` whose
`a:blip r:embed` names a synthetic 1×1 PNG the law adds to the package, and
`REQUIRED_SIBLING_LIMIT` counts the fillers a level generates rather than the
subject and the seeds placed at it, so the `a:graphic` is no longer crowded out
of a `wp:anchor` by the five siblings the schema also requires. The drawing then
takes the rebuild path and 21 pairs stop surviving:

- **16 `serialized-only-via-verbatim-replay`.** Ten are the wrap elements'
  own insets (`wp:wrapSquare@distT/B/L/R`, `wp:wrapTight@distL/R`,
  `wp:wrapThrough@distL/R`, `wp:wrapTopAndBottom@distT/B`). `ImageWrap` held
  one set of insets and the rebuild wrote it on `wp:anchor`, which is where
  OOXML reads them from when the wrap child states none — the effective value
  survived, the slot moved. Four are elements that carry nothing: an empty
  `<a:extLst/>`, an empty `<wp:cNvGraphicFramePr/>` (both documented as meaning
  the same as absence), and `wp:positionH/V`'s `<wp:align/>`, which the
  generator writes with no content and `ST_AlignH` admits no such value. The
  last two are `a:hlinkClick@r:id` and `a:hlinkHover@r:id`: the fixture's id
  names no hyperlink relationship, and `safeDocPrLinks` refuses to replay a
  link whose target it cannot resolve and check, which is the behaviour the
  link fix installed deliberately.
- **5 `present-with-a-different-value`.** `wp:wrapPolygon@edited` and the
  `@x`/`@y` of its `wp:start` and `wp:lineTo`. `serializeWrap` wrote a
  hard-coded rectangle — `edited="0"` and the four corners of a 21600-unit box —
  for every tight and through wrap, and nothing read the authored polygon.

Three more pairs — `wp:docPr@id`, `a:hlinkClick` and `a:hlinkHover` — became
measurable as `lost-in-the-editor-projection` and are fixed rather than
recorded: the drawing's id and both captured link elements are now carried on
the image node, so an edit that did not touch them hands their own bytes back.

### Two carriers for a wrap, and one outline

Fifteen of those pairs are now `modelled`, because the wrap is modelled as the
schema declares it rather than as one flat record.

`CT_WrapPath` is on `ImageWrap` in the units it is defined in — a 21600-unit box
over the drawing's extent, not EMU — read by `parseWrapElement` and carried
through the editor beside the wrap type on the image, shape and text-box nodes.
`CT_WrapTight` and `CT_WrapThrough` require the element, so a wrap that has no
outline still writes one; that rectangle is minted by the command that makes a
drawing tight or through, once, and is a value of the document from then on.
Minting it at the save instead would leave the editor holding a wrap the
document does not describe and would re-decide it on every save. A path with
fewer `wp:lineTo` than `CT_WrapPath` admits is written back as it stands: folio
is not the validator of its input, and substituting the drawing's full extent
would move text the source flows through the object.

The insets are two records, not one. `CT_Inline`, `CT_Anchor` and every
`EG_WrapType` member but `wp:wrapNone` declare their own set, and OOXML reads
the wrap child's where it states one and the drawing's otherwise — the fallback
chain the parser already had is only meaningful because they are two slots.
`ImageWrap.distanceSlots` records which element stated each inset beside the
value in force, and the rebuild writes each back there. A drawing whose slots
say nothing, or one an editor command has moved an inset on, states the value in
force on `wp:inline`/`wp:anchor`: the effective value is right either way, and
minting a wrap-child inset out of a moved value would not be.

The polygon the pairs above are measured on is now a legal one. `CT_WrapPath`
declares `minOccurs="2"` on `wp:lineTo` and the fixture used to carry a single
instance, because `fixture.ts` wrote one per required particle whatever its
minimum. The walk now tops each required particle up to the minimum it declares,
counting the instances the subject and the seeds already placed at that ordinal,
and caps the count at `REQUIRED_SIBLING_LIMIT` so no one particle can be what
makes a fixture unbounded. The cap does not bind: 22 particles in the graph ask
for more than one instance, the largest asks for three, and the only one
reachable from a rebuilt part's root is this `wp:lineTo`. The other 21 belong to
chart, theme and shape-geometry types a WordprocessingML part reaches only
through the `a:graphicData` payload the schema types as `xs:any`. Seventeen
fixtures carry a wrap polygon and each now writes two `wp:lineTo`; every one of
them reports what it reported before, so the measurement was right and only its
markup was not.

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
