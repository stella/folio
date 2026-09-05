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
`table-row-insert`, `table-row-delete`), for an agent that wants the summary
rather than the document. Every change carries the story it belongs to, so a
caller can tell a body edit from a footnote edit.

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

## How the alignment works

Aligning every paragraph in one pass cannot see structure: it pairs on text and
document order, so it will put a cell of one row opposite a cell of the next, or
a paragraph inside a table opposite one outside it. Rewriting such a pair in
place leaves the target's text in the wrong container. So the story is aligned
in three nested passes, each over things that can stand in for one another:

1. **Segments** — maximal runs of body text and of one table each, normalized to
   alternate, and paired by position.
2. **Rows** inside a paired table, by exact row text and then positionally, so a
   whole-row change stays whole and becomes `insertTableRow` / `deleteTableRow`.
3. **Cells** inside a paired row, by physical cell index.

Changed paragraphs go through `diffWordSegments` at apply time, so a redline
marks only the divergent words. Formatting-only differences are emitted as
`formatRange` operations and reported as `format`, never as a deletion and
reinsertion of identical text.

A comparison that finds no difference returns the base package as it arrived
rather than a re-serialization of it, so an unchanged document is handed back
byte for byte. The exception is a base that already carried tracked changes:
there the compared base is its accepted view, so the result is serialized.

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
- **Moves are reported, not represented.** The document carries a deletion at
  the source and an insertion at the destination; the change list keeps the
  relocation visible as `kind: "move"`. A relocated block needs at least three
  words to be recognized as a move, so boilerplate one-liners do not pair.
- **Tables cannot be created or destroyed.** The operation vocabulary has no
  "add a table", so a pair whose table count differs fails the round-trip check.
- **Empty cells are invisible.** A cell with no text carries no block, so a row
  whose cells are all empty is not seen as a row at all.
- **Column operations are out of scope.** A column added or removed reads as
  cell-level changes.
- **Row pairing degrades when a table's row count changes.** Rows match on exact
  text first and positionally after that, so a row deletion combined with cell
  edits can report per cell instead of as one row change. The result still
  accepts back to the target; it is just more granular than the edit was.
- **Numbering is not compared.** List renumbering that follows from an insertion
  or deletion is a property of the numbering definitions, not of block text, and
  no change is reported for it. A changed list level is invisible for the same
  reason and more seriously: `w:ilvl` moves and the comparison reports nothing
  at all, because the edit vocabulary has no paragraph-property operation to
  express it.
- **A split or a merge is overstated.** Both move a paragraph mark and no words,
  but are reported as a replace plus an insert or delete, so the half that did
  not change reads as newly written. `probes.test.ts` pins the current
  behaviour.

## Files

- `compare.ts` — the four stages (parse, align, apply, serialize), story
  pairing, and the round-trip self-check. `compareDocx` composes them.
- `plan.ts` — alignment and operation derivation. Pure.
- `formatting.ts` — the inline-formatting diff, shared with the redline
  generator.
- `reproducible-package.ts` — ZIP entry-date restamping.
- `scenario.ts` — the edit-script DSL the property tests build targets with.
- `../../scripts/compare.ts` — a manual runner for humans.
