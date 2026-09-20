---
"@stll/folio-core": minor
---

Write `w:tcW` from the width a cell states, not from the width the table resolved for it.

`TableCellAttrs.width` is the width the cell renders at: when the cell declares no `w:tcW`, the table resolves one from its grid and puts it there. The way back wrote that attr into `w:tcPr` unconditionally, so opening a document and saving it again gave every cell a preferred width its author never wrote.

`TableCellAttrs._authoredWidth` records what the cell itself states, as `_resolvedBorders` and `_resolvedMargins` already do for the border and the margin, and the save leg writes `w:tcW` from it alone. A command that moves a cell's width states one: `mergeTableCellAttrs` derives the record for every command that patches a cell, so a column resize writes exactly the cells it moved and a merge sums the widths its source cells stated rather than the widths the grid gave them.

The attr-schema version moves to 4. A version-3 snapshot states no `_authoredWidth`, and `_originalFormatting.width` is its record of which cells wrote a `w:tcW`, so the step backfills from it at the width the cell currently holds.
