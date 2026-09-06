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
  const { buffer, changes, unsupported } = result.value;
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
  otherwise writes the current time into every part it rewrites.

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
restores the base's. A mark is never written on the last paragraph of a table
cell: there is no sibling to join with, so the revision could not do what it
says.

Renumbering that FOLLOWS from an edit needs no change of its own: labels are
rendered from the numbering definitions rather than stored on the paragraphs,
so inserting a list item already renumbers the ones below it as-if-accepted,
and reporting them would bury the real edit. A definition that itself changed
is the opposite case — every label in the list moves and no block's text does
— and is reported as `numbering`.

A paragraph property that moved without any word moving — a list item demoted
a level, a paragraph restyled — is a `paragraph-format` change, written as
`w:pPrChange` with the complete previous property set so a reject restores it
the way Word does. The self-check's projection carries the style and the list
level alongside the text, so a redline that reproduces every word and leaves a
list item at the wrong level fails instead of passing.

`compareDocx` checks its own work before returning: accepting the generated
revisions must reproduce the target, table cell coordinates included. A
difference the operation vocabulary cannot express fails with
`CompareDocxRoundTripError` rather than returning a redline that reads
plausibly and is wrong.

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
  `kind: "move"`. Word also brackets each side with
  `w:moveFromRangeStart`/`End` and a shared `w:name`; those markers have no
  ProseMirror representation, so folio drops them on any edited paragraph and
  the comparison cannot produce them. A consumer reading the runs sees the
  move; one that groups multi-paragraph moves by range name does not.
- **Empty cells are invisible** (2026-09-06). A cell with no text carries no
  block, so a row whose cells are all empty is not seen as a row at all. The
  snapshot skips every empty textblock, and making it stop is a change to every
  block list in folio, not to the comparison.
- **Column operations are out of scope** (2026-09-06). A column added or
  removed reads as cell-level changes. `insertTableColumn` / `deleteTableColumn`
  exist in the operation vocabulary; nothing detects the difference yet, and
  detecting it reliably needs the empty cells above.
- **A table nested inside a cell cannot be added or removed** (2026-09-06). A
  whole table added or removed at document level is `table-insert` /
  `table-delete`; the same edit inside a cell would need `insertTable` to
  place a table in a cell rather than as a document-level peer.
- **A numbering definition is reported, not represented** (2026-09-06). A list
  whose format, level template or start changed is a `numbering` change, and
  the redline cannot carry it: OOXML has no tracked-change grammar for
  `numbering.xml` and Word does not track it either. Accepting the result
  therefore reproduces the target's words and keeps the base's numbering.

## Files

- `compare.ts` — the four stages (parse, align, apply, serialize), story
  pairing, and the round-trip self-check. `compareDocx` composes them.
- `plan.ts` — alignment and operation derivation. Pure.
- `formatting.ts` — the inline-formatting diff, shared with the redline
  generator.
- `reproducible-package.ts` — ZIP entry-date restamping.
- `scenario.ts` — the edit-script DSL the property tests build targets with.
- `plan.test.ts` — the judgement calls the corpus does not reach: the move
  similarity threshold, and row pairing when a table's row count changed.
- `probes.test.ts` — one labelled single mutation each, pinning what the
  change list SAYS rather than only that it round-trips.
- `../../scripts/compare.ts` — a manual runner for humans.
