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

### A transparent wrapper folio unwraps: `w:customXml` and `w:smartTag`

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
  *between* two children, it is *around* several, and a union member cannot
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

### The attribute remainder, and what it would take

`PreservedMarkup` once carried an ordered `attributes` list beside its
`children`, with a `modelsAttribute` predicate on the dispatcher and a
`serializePreservedAttributes` writer. Nothing in the product ever passed the
predicate, so nothing was ever kept: it was a flag with no effect, which
`AGENTS.md` bans, and it has been removed rather than left to look like
coverage. Its design is written here because the next pass should put it back
wired:

- The census charges 30-odd `@rsid*` pairs to `w:p`, `w:r`, `w:tr`, `w:tc`,
  `w:tbl`, `w:sectPr`, `w:pPr` and `w:rPr` as `never-parsed`. Word writes a
  revision-session id on nearly every one of those elements and folio rebuilds
  them without it, so a save rewrites the whole document's revision history.
- The predicate has to be a predicate, not a name set. folio resolves an
  attribute by namespace URI plus local name, and a remainder built by matching
  `"w:id"` textually keeps a second copy of a `w:id` a source spelled
  `altw:id`. Namespace declarations are not content: `captureVerbatimXml`
  rebinds what a captured fragment needs, and replaying a container's own
  bindings onto a rebuilt root fights the root's.
- The carrier is a field on the element's own model record
  (`Paragraph.preservedAttributes` and its siblings), not the child sink: an
  attribute has no position among children to keep.
- The editor leg is the boundary worth stating and testing. Attributes ride on
  the block's or run's preserved record and survive an ordinary round trip; an
  edit that rebuilds the element from scratch — splitting a paragraph, merging
  two runs — produces an element that never had those attributes, and it must
  not inherit a revision id from either neighbour. The rule is: the remainder
  follows the record, and a record the editor creates has none.
- `CAPTURE_SLOT_NAMES` in `laws.ts` has to name the field, or the survival law
  will clear nothing and the contract will call the pair `modelled`.

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
would hand back, so the *element* serializers run. For these parts the replay
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
