---
"@stll/folio-core": patch
---

Seven fidelity and review fixes:

- A paragraph whose `w:pPr` states only `w:ilvl` now keeps the `w:numId` its
  style supplies, so a demoted styled list paragraph stays numbered.
- Inserting a table row through a vertical merge extends the merge instead of
  splitting the grid.
- Adjacent paragraphs in a table cell that share a border definition draw one
  frame with the `w:between` rule, as they already do elsewhere.
- Every vertical `w:textDirection`, not only `btLr`, rotates its cell text.
- An abrupt-closing HTML comment (`<!-->`) no longer swallows the rest of a
  paste.
- A tracked replace whose two halves carry different timestamps is one review
  card again.
- Striking a selection leaves another author's existing deletion attributed to
  them.
