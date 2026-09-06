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
rather than the document.

## Determinism contract

The call is a pure function of its three arguments. Nothing reads a clock or a
random source:

- revision dates come from `options.timestamp`, which is required rather than
  defaulted so a caller cannot get an irreproducible package by omission;
- revision ids start one past the highest id the base package already carries;
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

`compareDocx` checks its own work before returning: accepting the generated
revisions must reproduce the target, table cell coordinates included. A
difference the operation vocabulary cannot express fails with
`CompareDocxRoundTripError` rather than returning a redline that reads
plausibly and is wrong.

## Limitations

- **Main story only.** Headers, footers, footnotes, and endnotes are reported in
  `unsupported`, not compared. So are parts present on one side only.
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
  no change is reported for it.

## Files

- `compare.ts` — the entry point, story pairing, and the round-trip self-check.
- `plan.ts` — alignment and operation derivation. Pure.
- `formatting.ts` — the inline-formatting diff, shared with the redline
  generator.
- `reproducible-package.ts` — ZIP entry-date restamping.
- `scenario.ts` — the edit-script DSL the property tests build targets with.
- `../../scripts/compare.ts` — a manual runner for humans.
