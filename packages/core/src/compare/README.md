# compare

`compareDocx(base, target, { author, timestamp })` returns the base `.docx`
carrying the tracked changes that turn it into the target, plus a JSON change
list describing them.

```ts
const result = await compareDocx(base, target, {
  author: "folio compare",
  timestamp: "2024-03-01T00:00:00.000Z",
});
if (result.isOk()) {
  const { buffer, changes, verification, unsupported } = result.value;
}
```

The buffer opens as ordinary revisions in any OOXML consumer. `changes` is a
discriminated union on `kind` (`insert`, `delete`, `replace`, `move`, `format`,
`table-insert`, `table-delete`, `table-row-insert`, `table-row-delete`,
`split`, `merge`, `paragraph-format`, `numbering`), for an agent that wants
the summary rather than the document. Every change carries the story it
belongs to, so a caller can tell a body edit from a footnote edit — except
`numbering`, which belongs to the package.

## Determinism contract

The call is a pure function of its three arguments. Nothing reads a clock or a
random source:

- revision dates come from `options.timestamp`, which is required rather than
  defaulted so a caller cannot get an irreproducible package by omission;
- revision ids start one past the highest id the base package already carries,
  and each story continues where the previous one stopped, so the ranges are
  disjoint without a guessed stride;
- `w14:paraId`s for paragraphs the comparison creates are derived from the
  stamp instead of `Math.random()`;
- ZIP entry dates are restamped from the same timestamp, because JSZip
  otherwise writes the current time into every part it rewrites;
- `dcterms:modified` in `docProps/core.xml` is restamped from it too, because
  the save otherwise dates the package from the second it happened to run and
  two runs then differ in that part alone.

Two runs over the same inputs therefore produce byte-identical buffers and
deeply equal change lists. `compare.property.test.ts` holds this as a property,
along with the round trip (accept-all yields the target, reject-all yields the
base), self-comparison, a churn bound, and the reporting of formatting-only and
move-only edits.

A relocated block is one `move` when it keeps at least three words and at
least 80% of its word tokens: a paragraph is usually edited on the way to its
new home, and a relocation that survives that is still a relocation. Below the
threshold it is a deletion and an unrelated insertion, which is what a reader
should be told when the text really did change that much.

## How the alignment works

Aligning every paragraph in one pass cannot see structure: it pairs on text and
document order, so it will put a cell of one row opposite a cell of the next, or
a paragraph inside a table opposite one outside it. Rewriting such a pair in
place leaves the target's text in the wrong container. So the story is aligned
in three nested passes, each over things that can stand in for one another:

1. **Segments** — maximal runs of body text and of one OUTERMOST table each,
   normalized to alternate. Paired by walking both sequences together: a table
   pairs with the table opposite it unless the next table on one side matches
   it better, which is what says the one in front of it was added or removed.
   Pairing by index instead cannot see a table appear or disappear — every
   later table shifts by one and each one's contents get rewritten into the
   next.
2. **Rows** inside a paired table, by similarity rather than by position: a
   row pairs with the row opposite it unless the next row on one side matches
   it better. Positional pairing is what made a deleted row plus a few cell
   edits report as a change in every row of the table, each row put opposite
   the one below it. Rows also carry their shape into the score, so two rows
   with the same words in a different number of cells are a replacement rather
   than an edit.
3. **Cells** inside a paired row, by physical cell index.

Changed paragraphs go through `diffWordSegments` at apply time, so a redline
marks only the divergent words. Formatting-only differences are emitted as
`formatRange` operations and reported as `format`, never as a deletion and
reinsertion of identical text.

An LCS on its own maximises matched characters, which on a rewritten sentence
means matching every stray "the" and comma it can reach and handing the reader
a dozen struck-through fragments interleaved with a dozen inserted ones. Three
rules pull it back: a match made only of separators is not a match, a
one-token match with changes on both sides of it is dropped into them, and a
paragraph whose surviving matches are too short for its length is replaced
whole. `options.granularity` cuts the redline at `"word"` (default) or
`"character"`.

`diffWordSegments` also takes case and whitespace normalization, which
`compareDocx` deliberately does not expose: a comparison that leaves a
difference unmarked does not accept back to the target.

A comparison that finds no difference returns the base package as it arrived
rather than a re-serialization of it, so an unchanged document is handed back
byte for byte. The exception is a base that already carried tracked changes:
there the compared base is its accepted view, so the result is serialized.

A split and a merge move a paragraph mark and no words, and are reported as
exactly that: `splitBlock` writes an inserted mark on the paragraph the break
now ends, `mergeBlockWithNext` a deleted one, and the change list says `split`
or `merge`. The space the break stands in for travels with the operation as
`separator`, so accepting reproduces the target's spacing and rejecting
restores the base's.

### Which paragraph mark changed

A paragraph mark belongs to the paragraph it ends: a paragraph added or removed
carries its own mark, inserted or deleted alongside its runs. That holds
everywhere but at the END of a container, which is the one place a mark cannot
say it went.

- **A mark with a paragraph after it resolves by joining.** Accepting a deleted
  mark closes the break; rejecting an inserted one does the same.
- **A mark with a table after it resolves by removing the paragraph.** There is
  nothing to join with, and once the resolution has taken the paragraph's words
  away, what is left is a paragraph whose break was resolved too, so it goes.
- **The mark that ENDS a container is never deleted.** A body, a table cell, a
  header or footer, a note and a text box each end with a paragraph, and that
  paragraph has none after it: a deleted mark there asks for a merge that
  cannot happen, and a consumer refuses the whole package rather than opening
  it. So a comparison that removes the paragraphs a container ends with runs
  the chain from the last SURVIVING paragraph forward: its mark goes, each
  removed paragraph's mark goes with it, and the container's final paragraph
  loses its words and keeps its mark as the carrier the merged text lands in.
  A paragraph's properties live on its mark, so the carrier's become the
  target's, written as `w:pPrChange` — that is bookkeeping for the merge and
  adds no entry to the change list, which says what it should say: the
  paragraphs were removed. "Its properties" means the ones the comparison
  compares at all, the paragraph style and the list level; a property outside
  that set is not read on either side, so the carrier keeps its own, and a
  reader accepting the redline sees the carrier's alignment or spacing rather
  than the surviving paragraph's. A carrier with no words to lose is not
  deleted at all: the removal is entirely the marks in front of it, and an
  operation that would write no revision is left out of the plan.
- **Paragraphs ADDED where the removed ones were land in the carrier.** The
  last of them is written into it as inserted runs before its kept mark; the
  rest become inserted paragraphs, marks and all, in front of it. So
  `[del A ¶del][del B][ins A'] ¶kept` accepts to `A'` and rejects to `A ¶ B`,
  and `[del A ¶del][del B][ins P1 ¶ins][ins P2] ¶kept` accepts to `P1 ¶ P2`.
  The mark count then works out on its own — one deleted for every removed
  paragraph but the carrier, one inserted for every added paragraph but the
  last — so the chain above is not rotated as well.
- **A carrier nothing precedes is reserved for the target's last paragraph.**
  A deleted mark joins two paragraphs of ONE container, so with nothing of that
  container in front of the carrier there is no chain a removal could run down:
  an empty body, or the blank paragraph a body that ends in a table has to
  carry because a table may not be a body's last child. The words the story
  ends with are then written into the carrier, and the alignment is told to
  hold the two last paragraphs out as a pair before it runs. That is a repair
  rather than a preference, and the alignment is what says whether it is
  needed: pairing the two ends where the chain does reach the carrier would
  trade a plain "this paragraph was removed" for a removal plus a rewrite that
  reads nothing like the edit.
- **An inserted mark at a container's end rotates the same way.** The break was
  ADDED, and rejecting an added break closes the paragraph it ends back over
  the NEXT one — which a container's last paragraph does not have, so the mark
  survives accepting everything and rejecting everything alike. The break
  therefore sits where it belongs: between the paragraph the run was appended
  after and the first appended one. That paragraph's mark is the inserted one,
  each appended paragraph but the last keeps an inserted mark of its own, and
  the paragraph the container now ends with takes the free mark, recording the
  other's properties as `w:pPrChange`. Only which paragraph is left markless
  changes, so `[A][B ¶ins][C ¶kept]` accepts to `A ¶ B ¶ C` and rejects to
  `A ¶ B`. The rotation reaches across paragraphs and no further: a table among
  the appended blocks stops it, and the words the target ends with are written
  into the base's own last paragraph instead.

The serialize stage refuses a package whose final paragraph mark carries a
revision in either direction, with `CompareDocxFinalParagraphMarkError` naming
the container and the paragraph. That one is fatal under `onUnverified:
"emit"` too: unlike an unproven redline there is no partial result worth
handing back, because the file does not open, or opens carrying a revision no
reader can clear. A cell of a row the package is DELETING is the exception the
format asks for: `w:trPr/w:del` plus a deletion on every mark its cells end
with is how a removed row is written, and those marks leave with the row.

Which paragraph ends a container is settled only when the whole batch is, so
the rotation runs once over the finished document rather than at each
insertion: an insertion that looked final is not one after the next operation
writes a table past it.

A body and every table cell end with a paragraph: a table may never be a
container's last child. A comparison that adds a table at the end therefore
adds the paragraph after it too, and that paragraph is an insertion like any
other.

A relocated paragraph's break is `w:moveFrom` at the source and `w:moveTo` at
the destination. They resolve exactly as a deletion and an insertion do; the
kinds are what tell a reader the two ends belong together rather than being an
unrelated removal and addition.

Renumbering that FOLLOWS from an edit needs no change of its own: labels are
rendered from the numbering definitions rather than stored on the paragraphs,
so inserting a list item already renumbers the ones below it as-if-accepted,
and reporting them would bury the real edit. A definition that itself changed
is the opposite case — every label in the list moves and no block's text does
— and is reported as `numbering`.

A paragraph property that moved without any word moving — a list item demoted
a level, a paragraph restyled — is a `paragraph-format` change, written as
`w:pPrChange` with the complete previous property set, which is what a reject
restores. The self-check's projection carries the style and the list
level alongside the text, so a redline that reproduces every word and leaves a
list item at the wrong level fails instead of passing.

## Verification

`compareDocx` checks its own work before returning, in both directions:
accepting the generated revisions must reproduce the target, and rejecting
them must reproduce the base they were written against — table cell
coordinates and the tables' own properties included.

A difference the operation vocabulary cannot express fails with
`CompareDocxRoundTripError` rather than returning a redline that reads
plausibly and is wrong. That is the default, and it is the right default: a
reader cannot tell a redline that lost something from one that did not.

`onUnverified: "emit"` asks for the other trade. The call then returns the best
redline it could build and a `verification` that names what it could not prove:

```ts
const result = await compareDocx(base, target, {
  author: "folio compare",
  timestamp: "2024-03-01T00:00:00.000Z",
  onUnverified: "emit",
});
if (result.isOk() && result.value.verification.status === "unverified") {
  for (const { invariant, cause, story, detail } of result.value.verification.failures) {
    // invariant: "accept-reproduces-target" | "reject-reproduces-base"
    // cause: which field of the block projection diverged
  }
}
```

`verification` is on every successful result, so a caller that never passes the
option still sees `{ status: "verified" }` and can assert on it. `cause` is one
of `invisible-structure`, `block-count`, `container`, `table-geometry`,
`style`, `list-level`, `whitespace`, `text` — the projection field that
diverged, which is what names the part of the pipeline that lost the
difference. `table-geometry` is the one that no block carries: a second
projection reads each table's `w:tblPr`, `w:trPr` and `w:tcPr` so a redline
that reproduces every word and none of the widths, spans, merges, shading or
borders fails instead of passing. `invisible-structure` is the
one cause that is not a lost difference: every block is present, in order, at
coordinates the block model cannot reach, because the snapshot carries no block
for an empty paragraph.

Every `detail` is structural — counts, offsets, container kinds — and carries no
phrase of either document, so it is safe to log, report, or quote.

A parse, apply or serialize failure is still an error under either setting:
there is no redline to emit. So is any revision on a container's final
paragraph mark, for the same reason at the other end — the bytes would be
written, and no consumer would open them, or would open them carrying a
revision neither accepting nor rejecting everything can clear.

A SKIPPED operation is not automatically one of those. An operation the applier
had nothing to write for, or could not write in that place — a paragraph inside
a text box or a content control, where a paragraph mark has nowhere to go —
leaves the redline standing, and whether anything was lost by it is the
question the round trip already answers. A skip that says the plan did not
match the document it was planned against is the other kind: the operations
came from that very snapshot moments earlier, so nothing should have moved
under them, and the call fails with `CompareDocxApplyError`.

## Inputs that already carry tracked changes

Both sides are compared **as accepted**. Either document may arrive carrying
someone else's unresolved revisions, and there are only two things a comparison
can do with them: layer its own marks on top, or resolve them first.

Layering does not survive contact with a reader. The package would hold two
redlines by two authors with no way to tell which belongs to the comparison,
and rejecting everything would land on a third document neither side wrote.
Resolving to the accepted view states one question instead — how does the base
as it stands differ from the target as it stands — and leaves the answer as the
only redline in the package. The round trip is then exact: rejecting returns
the base's accepted view, accepting returns the target's.

A caller who wants the earlier revisions preserved should resolve them
deliberately before comparing.

## Tables

A table the comparison adds is the target's table, not a grid of its cell
texts: `w:tblPr` with its style, width, indent, justification, borders, cell
margins, layout and look; `w:tblGrid` with the target's column widths; each
row's `w:trPr`; each cell's `w:tcPr` with `w:tcW`, `w:gridSpan`, `w:vMerge`,
`w:shd`, `w:tcBorders`, `w:tcMar` and `w:vAlign`; cell paragraphs with their
own properties and runs; and tables nested in cells, recursively. The operation
vocabulary is unchanged — `insertTable` and `insertTableRow` still describe
their content as cell texts, because that is what a caller writing a table from
nothing has — and the node travels beside the batch instead, since a document
node is not JSON.

A table crossing from one package into the other cannot bring what only the
first package can resolve, so a copied table drops hyperlink, note and comment
marks, paragraph identities, and any inline content that names a relationship.
Everything that describes the table itself travels.

A table or row the comparison removes keeps every property it had, under the
deletion marks: the row carries `w:trPr/w:del` and every run inside it carries
`w:del`, so rejecting restores the table exactly and a consumer that reads only
one of the two marks still resolves the deletion.

A table that stayed in place while its properties changed moves no block, so no
block operation carries the difference. Each paired table, row and cell whose
properties differ is rewritten to the target's and records the previous set as
`w:tblPrChange` / `w:trPrChange` / `w:tcPrChange`, which is what a reject
restores.

## Limitations

`benchmarks/compare` measures each of these; `benchmarks/compare/RESULTS.md`
carries the current numbers and the failing cases.

- **Parts present on one side only are reported, not compared** (2026-09-05).
  Creating or removing a header, footer or note part is not a text edit, so
  such a story is listed in `unsupported`. Every story present on both sides is
  compared: main, headers, footers, footnotes and endnotes, each with its own
  revision id range so no two stories claim the same `w:id`.
- **A text box is compared as body text** (2026-09-05). Its paragraphs are part
  of the main story, so their text is compared, but the round-trip self-check
  tags only the enclosing table cell — an insertion that landed inside a box
  instead of beside it would not be caught by it.
- **A move's range markers are not written** (2026-09-06). The relocation
  itself is in the document: the deletion at the source is `w:moveFrom`, the
  insertion at the destination `w:moveTo`, and the change list reports
  `kind: "move"`. The format also brackets each side with
  `w:moveFromRangeStart`/`End`, `w:moveToRangeStart`/`End` and a shared
  `w:name`; those markers have no ProseMirror representation, so folio
  drops them on any edited paragraph and the comparison cannot produce them. A
  consumer reading the runs sees the move; one that groups multi-paragraph
  moves by range name does not.
- **Ambiguous column operations are refused** (2026-09-06). The comparison
  emits `table-column-insert` / `table-column-delete` only when the unchanged
  columns give one exact grid alignment. Grid coordinates and spans come from
  `TableMap`; the snapshot's `cellIndex` remains a physical row-child index.
  Repeated empty columns, edits that cut a horizontal span, and target columns
  containing vertical spans stay unverified rather than being guessed.
- **A table nested inside a cell cannot be added or removed** (2026-09-06). A
  whole table added or removed at document level is `table-insert` /
  `table-delete`; the same edit inside a cell would need `insertTable` to
  place a table in a cell rather than as a document-level peer. A nested table
  travels with the table that holds it, so one added or removed alongside its
  parent keeps its own grid and properties.
- **A paired table's grid is not moved** (2026-09-07). `w:tblGrid` changes are
  recorded with `w:tblGridChange`, which the editable model does not carry, so
  a table whose columns kept their text and changed their widths keeps the
  base's. A table the comparison ADDS carries the target's grid, because there
  is no previous one to record.
- **A property a table style resolves is not moved** (2026-09-07). A change
  element stores the complete previous property set, and rejecting it rebuilds
  the live properties from that record alone. Where the base's effective
  properties cannot be rebuilt from what would be written for it — a value a
  style resolves that no `w:tcPr` of the document states — the difference is
  left alone rather than written as a revision that rejects to a third
  document.
- **`colspan` and `rowspan` are not moved on a paired cell** (2026-09-07). They
  shape the table's map, and changing one without restructuring the rows around
  it leaves the map inconsistent with its own grid. A span that changed is a
  row or column edit, not a property change.
- **A numbering definition is reported, not represented** (2026-09-06). A list
  whose format, level template or start changed is a `numbering` change, and
  the redline cannot carry it: OOXML has no tracked-change grammar for
  `numbering.xml` at all. Accepting the result
  therefore reproduces the target's words and keeps the base's numbering.

## Files

- `compare.ts` — the four stages (parse, align, apply, serialize), story
  pairing, and the round-trip self-check. `compareDocx` composes them.
- `plan.ts` — alignment and operation derivation, including which table each
  `insertTable` / `insertTableRow` should place and which cells were paired.
  Pure.
- `../ai-edits/table-template.ts` — copying a table out of one package into the
  other: what travels, what cannot, and how an insertion is stamped.
- `../ai-edits/table-geometry.ts` — the table-property projection the
  self-check compares, and the matching that moves a paired table's properties.
- `verification.ts` — the round-trip verdict: the invariants, the causes, and
  the safe-to-quote detail each failure carries, plus the structural guard on a
  container's final paragraph mark. Pure.
- `formatting.ts` — the inline-formatting diff, shared with the redline
  generator.
- `reproducible-package.ts` — the clocks outside the document body: ZIP entry
  dates and `dcterms:modified`.
- `scenario.ts` — the edit-script DSL the property tests build targets with.
- `plan.test.ts` — the judgement calls the corpus does not reach: the move
  similarity threshold, and row pairing when a table's row count changed.
- `probes.test.ts` — one labelled single mutation each, pinning what the
  change list SAYS rather than only that it round-trips.
- `../../scripts/compare.ts` — a manual runner for humans.
