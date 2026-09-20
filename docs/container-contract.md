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
  back, save, and find the subject in the saved part under the chain the fixture
  wrote it at, as many times as the fixture wrote it, with an equal value.
- **L3 editor** — the same through `toProseDoc`/`fromProseDoc`. L3 projects
  with reuse declined; a pair that survives only because the merge reused its
  base block is not a projection survival. The law projects through one helper,
  `scripts/lib/container-survival/projection.ts`, so there is a single place
  the decline is stated, and a test binds that call to the conversion's own
  signature: the option cannot be added without the law taking it.
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

### What counts as the same name

Those three are about a value. A fourth equality is about a name, and it is the
same idea one level up: a document may spell a handful of things two ways, folio
writes one of them, and a probe that matches the subject by its authored
spelling — `<w:start`, ` w:start="` — reports the rename as a loss although
nothing was lost. `CANONICAL_SPELLINGS`
(`scripts/lib/container-survival/canonicalSpellings.ts`) is the list of those
equivalences, and the probe accepts a canonical spelling **only** for a subject
an entry names:

| A document may write                            | folio writes         | Why they are the same thing                                                                                                                                                                  |
| ----------------------------------------------- | -------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `w:start`/`w:end` of `CT_Border`, `CT_TblWidth` | `w:left`/`w:right`   | The logical-direction name of the physical one. The Transitional schema declares both as optional siblings of the same type, folio's readers take either, and its model holds one direction. |
| `w:ind@w:start`/`@w:end`                        | `@w:left`/`@w:right` | The same rename, as an attribute.                                                                                                                                                            |
| `CT_OnOff@w:val` = `true`, `1`, `on`            | the attribute absent | A `CT_OnOff` element with no `w:val` is the on state. The off state keeps its `w:val="0"`, so only the on spellings are listed.                                                              |
| `CT_VMerge@w:val` = `continue`                  | the attribute absent | An omitted `w:val` is a continuation cell (`vmerge-absent-means-continue`).                                                                                                                  |
| `CT_TabStop@w:leader` = `none`                  | the attribute absent | An omitted leader draws nothing.                                                                                                                                                             |

The last three are values, and they reach the value half of the probe only: an
absent attribute means what the fixture wrote, an element that did not come back
at all is still absent, and `w:ind@w:startChars`, which folio really does drop,
still reads as lost.

The table is hand-written, like the fixture-realism tables below and unlike
everything else the law reads, and it fails in the worse direction: a wrong
entry reads a real loss as a survival. It is also total against nothing, and
cannot be — there is no schema-declared class of names folio canonicalises to be
total over, because the schema declares `w:start` and `w:left` as two ordinary
siblings and records no default for any of the three attributes. So the guard
runs the other way: every entry must be exercised by a pair the census generates
(`container-survival-canonical.test.ts`), which makes the table shrink-only, and
every entry carries the rule it rests on and the serializer line that performs
it, read back from that file so a moved line fails rather than rots.

### Where the law looks, and how many it wants

The probe is the law. It used to ask whether the subject's name appeared
anywhere in the saved part, and both halves of that question were wrong.

- **Anywhere.** A pair is (container, child). The `w:pgSz` of the section
  snapshot a `w:sectPrChange` holds is not the live section's `w:pgSz`, and a
  part that lost the first still contains the second. The same goes for every
  wrapper folio unwraps: a `w:smartTag`'s runs are spliced into the paragraph,
  a `w:bdo`'s content reaches the editor without the wrapper, a row-level
  `w:bookmarkStart` was re-anchored inside a cell's paragraph, and a marker
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

One of these groups is worth a decision rather than a fix. The marker hoist is
deliberate: document order is unchanged and the wrapper splits in two, which is
the behaviour the section below describes as harmless. It is recorded as
`dropped (parsedNotSerialized)` because that is where the markup ends up, and
the alternative is a probe that knows which relocations folio meant — the
leniency this change removed. The bookmark re-anchoring was the same shape and
was not harmless; the next section is what it cost and what it took to stop.

### A bookmark keeps the container it was written in

Eight of those ten re-anchored pairs were a bookmark, and they are now
`modelled`. Along with the header and footer pairs the probe had recorded as
`serialized-only-via-verbatim-replay`, and the table's, fourteen pairs moved —
every one of them toward the stronger disposition, none the other way.

**The markup is what Word writes, not an edge case.** A scan of the 5335
packages in the public corpus finds 1050 bookmark markers standing as a direct
child of `w:body` (658), a block `w:sdtContent` (162), `w:tr` (95), `w:tbl`
(77) or `w:tc` (58), across 379 files. Sixteen pairs open _and_ close on a
`w:tr`: a bookmark over whole rows. `_GoBack` alone accounts for much of the
content-control half.

**What was wrong was the extent, not the survival.** The marker reached the
saved part either way, which is why only a probe that asks _where_ could see
it. Re-anchored into the neighbouring paragraph, a range that covered two
paragraphs came back covering neither — both halves landed inside the same
paragraph at the same ordinal, a zero-width bookmark — and a row-spanning range
came back inside one cell. A `REF` field or a link resolving either one then
covers the wrong text.

**The carrier is typed at every level, and that is the load-bearing decision.**
The obvious cheap fix is the verbatim sink: the row and the table already have
one, and it would have held the marker exactly where it stood. It is the wrong
answer here, and the corpus says why. Of the 730 pairs with a block-anchored
half, 394 have their _other_ half inside a paragraph — `w:tr` open to `w:p`
close (62), `w:p` open to `w:tbl` close (51), and so on — so the two halves are
routinely at different levels. folio pairs a start with its end over the model,
in `collectPairedBookmarkIds` on the way in and in the boundary integrity pass
on the way back, and neither can see inside a capture. A half kept as bytes
leaves the other half unpaired, and an unpaired boundary is deleted. The
bookmark would not merely move; it would go.

So the marker is a member of `BlockContent` wherever the container models a
sequence of blocks — a body, a cell, a block content control — which is the
union case the section above prefers, with no index to keep honest. A row and a
table model one kind of child, so theirs is `TableRow.bookmarks` /
`Table.bookmarks`: an index among the cells or rows, beside the verbatim sink
rather than inside it. Same position mechanism, different contents, and the
reason is exactly that one of them has to be readable.

**The editor leg.** `blockBookmarkBoundary` is the block twin of the inline
`bookmarkBoundary` atom; ProseMirror decides inline or block per node type, so
the two levels cannot be one node, and everything else about them — the
attribute spec, the DOM, the reader — is shared so they cannot describe a
bookmark differently. A row's and a table's markers ride the node's own attrs
by reference, the way an attribute remainder does, and the integrity pass reads
all four carriers, so a pair spanning two levels stays whole while either half
is being edited.

Two pairs of the original ten are not this fix and stay
`dropped (parsedNotSerialized)`: `CT_SdtContentRun`'s, which are the _inline_
content control's. `INLINE_SDT_CONTENT` lifts a bookmark out of `w:sdtContent`
as a sibling of the `w:sdt`, and closing that is the inline widening
`inlineWrapperContent.ts` already owns.

### What seeing the canonical form moved

Teaching the probe the spellings above moved 72 pairs, every one of them to
`modelled`, and nothing in the other direction. They are a correction to the
measurement, not a change to anything folio writes.

| Pairs | Was                                       | What the probe was missing                                                                                                                                                          |
| ----- | ----------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 35    | `serialized-only-via-verbatim-replay`     | `CT_OnOff@w:val` on 35 elements: the fixture wrote `w:val="true"` and folio wrote the bare element.                                                                                 |
| 5     | `never-parsed` (4), `replay-rejected` (1) | The same omission on `w:titlePg`, `w:formProt`, `w:noEndnote`, `w:rtlGutter` and `w:showingPlcHdr`, where the missing attribute became a verdict about the element that carries it. |
| 20    | `serialized-only-via-verbatim-replay`     | The attributes of a `w:start`/`w:end` that came back as `w:left`/`w:right`: eight of `CT_Border`'s nine, and both of `CT_TblWidth`'s.                                               |
| 8     | `serialized-only-via-verbatim-replay`     | Those elements themselves, in `w:tblBorders`, `w:tcBorders`, `w:tblCellMar` and `w:tcMar`.                                                                                          |
| 4     | `serialized-only-via-verbatim-replay`     | `w:ind@w:start`/`@w:end`, `w:vMerge@w:val="continue"` and `w:tab@w:leader="none"`.                                                                                                  |

`CT_Border@w:themeColor` on `w:start` and `w:end` stayed lost, which is the
control: the element is found under its canonical name, and the attribute is not
on it, because nothing in the model holds one.

The value sweep grew by 21, and that is a defect the correction uncovered rather
than one it caused. A slot that loses its representative value is not swept for
the rest of its type, so seven `CT_OnOff` elements — `w:cantSplit`,
`w:tblHeader`, `w:noWrap`, `w:tcFitText`, `w:hidden`, `w:showingPlcHdr` and
`w:specVanish` — had never been measured on an off value. Their serializers
write the bare element for the on state and nothing at all for the off one, so
an explicit `w:val="0"`, which is what overrides an inherited setting, is
dropped. Three spellings of off on seven elements is the 21.

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

**An attribute pair is measured twice: alone, and beside one the element
models.** A fixture that states one attribute at a time cannot see a reader
that decides an element whole, and it reports every attribute of such an
element as surviving — `<w:ind w:leftChars="100"/>` is kept because nothing
was taken from it, and the `<w:ind w:left="720" w:leftChars="100"/>` a
document carries is not. So `modelledCompanionFor` states a second attribute
on the same element and the pair survives only when it survives both runs.
Which attribute counts as modelled is the model's answer rather than the
census's: the companion is the first different field of the same
`PROPERTY_ELEMENT_ATTRIBUTES` table the reader computes its remainder from, so
the two cannot disagree about the word or pair two spellings of one field.
Nothing is added where the element
declares a _required_ modelled attribute — `CT_TabStop`'s `w:val` and
`w:pos`, `CT_Shd`'s `w:val` — because the ordinary fixture already states
those, and the second run is skipped where the first already lost.

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
lost with the two row wrappers, move to `dropped (editorProjection)` — all but
the two bookmark pairs, which `Table.bookmarks` carries on the table node's
attrs for the reason the row section now gives.

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
about the inner element's _parent_ rather than about the part containing a
name, because the flattening failure passes every probe that only counts.

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
`dropped (neverParsed)` to `dropped (editorProjection)`. The two bookmark pairs
are the exception, and they show what the rest would cost: a marker folio can
read does not need a node, because `TableRow.bookmarks` rides the row node's
attrs by reference. That only works for markup the model holds — bytes on an
attribute are bytes the editor still cannot place — which is why the bookmark
is typed and the sink is not. That is not a lateral move. `neverParsed` says folio never read the markup and a document that is
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
  the generated list is _a_ valid order rather than the only one — which is
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
  anti-vacuity guard asked that _some_ member of a row refuse a reversal, which
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
the width. `TableCellAttrs.width` is the width the cell _renders_ at, which the
table resolves from its grid when the cell declares no `w:tcW`, and the way back
wrote it into `w:tcPr` unconditionally, so a cell that stated none acquired one.
`TableCellAttrs._authoredWidth` is the record of what the cell states, as
`_resolvedBorders` and `_resolvedMargins` are for the border and the margin, and
the save leg writes `w:tcW` from it alone. A command that moves a cell's width
states one: `mergeTableCellAttrs` derives the record for every command that
patches a cell, so a resize writes exactly the cells it moved.

### The same property set with four owners

`w:rPr` is another property set and the first with more than one owner.
`CT_RPr` is `EG_RPrBase` plus `w:rPrChange`; `CT_ParaRPr` opens that with
`EG_ParaRPrTrackChanges`; and `CT_RPrOriginal` and `CT_ParaRPrOriginal` are the
snapshots inside either one's revision. One generated row covers all four, and
`TextFormatting.preserved` is the one sink.

It joins the sequence rows for a weaker reason than the other two, and the
difference is worth stating rather than inheriting. `CT_TblPrBase` is a real
`xsd:sequence` of distinct names, so a `w:tblPr` out of order is invalid and
`validateOoxmlPart` says so. `EG_RPrBase` is an `xsd:choice` referenced
`maxOccurs="unbounded"`: a `w:rPr` in any order is valid, a repeated child is
valid, and the validator reports neither. What the generated order buys here is
one canonical form — the one Word writes — from folio's two `w:rPr` writers,
which is why the order lives in `@stll/docx-core` where both can read it. The
repeat is answered by the reader instead: the last statement of a property
wins, and the statements it beat are not written back, so a canonical-order
writer cannot invert them. See `docs/reserved-values.md`.

Three further things about the row are decisions rather than consequences.

- **The owners differ only in which children a _sibling_ record claims.** A
  run's `w:rPrChange` is read into `Run.propertyChanges`; the paragraph mark's
  revision and its `w:specVanish` are read into the paragraph's own record.
  A child two readers both take is written twice; one neither takes is lost.
  So the call site names its owner and the map overrides exactly those
  children, the way the hyperlink's two callers already do.
- **`ownedElsewhere` is a claim about the whole child, and a record that holds
  only part of one may not make it.** `ParagraphFormatting.runInWithNext` is
  on-or-absent, so it has nowhere to put `<w:specVanish w:val="0"/>`, the value
  that cancels a style's run-in heading. That child's disposition is a handler
  answering with what the record took, not a name in a map.
- **The editor leg stops at the run, and the reason is the one the attribute
  remainder already gave.** The paragraph mark's properties ride
  `ParagraphAttrs._originalFormatting.runProperties`, so its sink and its
  `w:rPrChange` reach the editor and come back. A run has no such record: a run
  is text plus marks, so `r|CT_R`'s own `w:rPr` sink survives a save and not a
  round trip, and `rPr|CT_RPr`'s eleven pairs are `editorProjection` rather than
  `containerNotKept`. Giving a run one is the same separate record with the same
  grouping rules the remainder needs.

A merge is where this sink differs from the other property sets. Run formatting
is resolved: a style, the paragraph mark and the run each have their say, and
`preserved` is not a value that resolves. It is the bytes one element held, so
`mergeTextFormatting` drops it on every path out rather than writing a style's
markup into every run below it.

**An empty property set is not an absent one**, and `w:rPr` is where that rule
costs the most. The element is optional wherever it appears, so the reader
answers `undefined` only for an owner that carried none, and the writer writes
`<w:rPr/>` exactly when the record is present — the decision `w:tblPrEx`,
`w:trPr` and `w:tcPr` already made. The corpus argues for it at every owner
(5121 empty ones on a run, 3890 on the paragraph mark, 2419 on a style, 674 on
a numbering level, 31 inside a `w:rPrChange`, Word among the producers of
each), and the paragraph mark is the sharpest case: an empty one states no
formatting, which is the whole point, but it is the slot the mark's
`w:rPrChange` and its `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` live in, so
writing one is pure presence. The emission carries three answers rather than
two for the same reason — `undefined` for a mark with no property set, `""`
for one with an empty set — because a caller handed `""` for both puts the
presence back on the floor.

Three pairs move with it, and one of the three is the census rather than folio.
`pPr|CT_PPr/rPr` goes from `dropped (replayOnly)` to `modelled`, and
`r|CT_R/rPr` from `dropped (parsedNotSerialized)` to `dropped
(editorProjection)` — the model holds it and a save writes it; the editor has
no run record to carry it, which is the boundary the attribute remainder
already names. `rPrChange|CT_ParaRPrChange/rPr` reads as `modelled` and is
not: the paragraph mark's `w:rPrChange` is still a capture in the sink, and
what the carrier probe finds with the sink cleared is the paragraph mark's own
`<w:rPr/>`. The law asks whether the subject is _somewhere_ in the part with an
equal value, and one empty element is equal to another. A pair whose fixture
value is an empty element cannot be told from a sibling of the same name until
the law compares positions.

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

### A property set: the same sink, counted differently

`w:pPr` is the last of the property sets to get a dispatcher, and it is where
the sink's `index` stops meaning what it means everywhere else. Three things
about it are decisions rather than consequences.

- **The ordinal is the schema's, not a count of siblings.** `CT_PPrBase` is an
  `xsd:sequence` of thirty-three distinct optional names, so a child's place is
  a property of its name. A count of modelled siblings is a mirror of whichever
  properties folio models today: model one more and every capture recorded
  before it slides one place. `sequencePositions` records the declared ordinal
  instead, and `serializeSequenceChildren` merges the modelled children and the
  captures by it. An undeclared name has no place in the sequence, so it takes
  the place of the last declared child before it.
- **The order is generated, and it lives in the lower package.** The order
  belongs to the model rather than to any one serializer, and four writers
  produce the element: a paragraph, a style's `CT_PPrGeneral`, a numbering
  level's, and the `CT_PPrBase` snapshot inside `w:pPrChange`. The generated
  table is `SEQUENCE_CHILDREN` in `@stll/docx-core/schema`, folio-core's
  declared-child table spreads it in, and one writer emits the element for all
  four — so the set a handler map is total over and the order a serializer
  writes cannot disagree.
- **A handler may refuse.** `<w:spacing/>` states no spacing and
  `<w:jc w:val="end"/>` states a value the reader's enumeration does not admit.
  Neither can be decided by name — a map keyed by name can list the names a
  reader has never heard of, never the values it will refuse — so the handler
  answers `CAPTURE` when it took nothing and the bytes are kept.

The sink rides the editor as an `original-only` field of the paragraph's
formatting attrs, the rule the attribute remainder states one level up: the
remainder follows the record, and a record the editor creates has none. And
the cascade drops it. A style's captured bytes are not a paragraph's direct
formatting: `mergeParagraphFormatting` strips `preserved` from both tiers, so
the same markup cannot be written at two of them and an inherited value cannot
come back outranking the tier it came from.

The shared property-set reader moves 66 pairs to `captured-verbatim`: 50 that
previously survived only through replay, plus 16 whose property-set container
was not kept. They cover the unmodelled `w:pPr` children and the attributes
carried by those children, including `w:cnfStyle`'s twelve flags.

### The attribute half of a property element

A handler answers about the element, so the sink's decision was
whole-or-nothing: `<w:ind w:leftChars="100"/>` was kept entire because the
reader took nothing from it, and `<w:ind w:left="720" w:leftChars="100"/>` —
which is what a document carries — was modelled and lost the character unit.
This is the attribute remainder again, one level down from `w:p`'s `w:rsid*`
attributes, and four things about it are decisions.

- **The remainder rides the record that holds the element's modelled fields.**
  For `w:framePr`, `w:tab`, a `w:pBdr` side and `w:shd` that record is the
  element's own, so each gains a `preservedAttributes`. `w:ind` and
  `w:spacing` were flattened into `ParagraphFormatting`, so that is their
  record and it carries one remainder per flattened element —
  `indentPreservedAttributes` and `spacingPreservedAttributes` — rather than
  one for the set. Same rule, applied to where the model actually put the
  fields.
- **The predicate is derived from the model, not written beside the reader.**
  `propertyElementAttributes.ts` holds one table per record, each
  `as const satisfies ModelledAttributes<…>` over the record's own fields, so
  a field added without an attribute to name does not compile. A field may
  name several: `indentLeft` is filled from `w:left` or from the Strict
  `w:start`, and both have to be out of the remainder or a save writes the
  same indent twice under two spellings. `readAttributeBag` takes the table
  and nothing else, so no call site can state a set of its own.
- **`w:shd` and `w:framePr` compute a remainder only when they are modelled at
  all.** An element folio takes nothing from goes to the dispatcher's sink and
  is kept whole; a remainder as well would write its attributes twice.
- **The census had to change with it, or the fix could not be measured.** A
  pair stated alone says nothing about a whole-or-nothing reader. See
  [Fixture realism](#fixture-realism).

Four more losses close in the shared shading reader: `w:themeTint`,
`w:themeShade`, `w:themeFillTint`, and `w:themeFillShade` now remain modelled
even when the source states a modifier without its base theme slot. Ten
attributes that a one-at-a-time census reported as surviving are now measured
beside a modelled sibling and survive that too: `w:ind`'s six character units,
`w:spacing`'s two line counts, and `w:framePr`'s `w:hRule` and `w:anchorLock`.
Together, the property-set sink and shading modifiers move 70 pairs out of the
dropped set.

An empty prior property set still carries a change record: keeping that
`w:pPrChange` closes five more pairs, covering the wrapper, its three tracked
change attributes, and its nested `w:pPr`.

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

The effect extent is two records for the same reason, and unlike the insets they
are two values. `CT_Inline` and `CT_Anchor` declare a `wp:effectExtent`, and so
do `CT_WrapSquare` and `CT_WrapTopBottom`: the drawing's is the object's own
effect reservation, which `Image.padding` holds, and the wrap child's is the
reservation the text flow is computed against. folio read only the drawing's and
wrote it back there, so the wrap child's came back on the anchor and its two
pairs read as `serialized-only-via-verbatim-replay`.
`ImageWrap.effectExtentSlots` records which element stated which, and
`resolveEffectExtents` writes each back there while the drawing's is unmoved.
Once an editor has resized the drawing's, the rebuild states that one alone: a
wrap reservation computed against a shape that is no longer there would flow
text around nothing. An all-zero reservation and no reservation are the same
document, because `CT_EffectExtent` requires all four sides and a rebuild writes
zero for a side the record holds none for — so the drawing's slot drops an
all-zero element and the wrap child's keeps it, where an absent one means the
drawing's instead. A shape and a text box hold no reservation of their own, and
had `l="0" t="0" r="0" b="0"` written on every rebuild; they now keep the one
they were authored with.

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

### An element that states nothing

`w:sectPrChange` holds a `CT_SectPrBase`: `CT_SectPr` without the header and
footer references and without a change record of its own. Everything else it
declares the live section declares too, and one reader and one serializer answer
for both, so the two cannot disagree — which is what made
`sectPr|CT_SectPrBase/lnNumType` and `/pgBorders` worth the look, because they
did.

The cause was not the snapshot. `CT_LineNumber` and `CT_PageBorders` declare
every attribute and every child optional, so `<w:lnNumType/>` and
`<w:pgBorders/>` are legal elements saying a section is line numbered, or
bordered, with the defaults. Both serializers wrote nothing for a record holding
nothing, and nothing else in the repository writes those fields: the record is
there because the parser read the element. `serializeDocGrid`, in the same file,
had already decided the other way.

The two places differ in how loudly the loss is made. On a live section the
record then holds a field the serializer did not write, `w:sectPr` comes out as
the empty string, and the package fidelity guard refuses the save. Inside a
change record `serializeSectionPropertyChange` reads that empty string as
nothing to write and puts back `<w:sectPr/>`, so the section a reviewer would
restore comes back blank and nothing says so. Asking _where_ is what separated
them: the probe looks under `w:sectPrChange/w:sectPr`, not for a `w:lnNumType`
anywhere in the part.

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
