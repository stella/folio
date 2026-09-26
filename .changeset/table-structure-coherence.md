---
"@stll/folio-core": patch
---

Deleting a table column directly and accepting the same deletion tracked now leave the same table. Where the deleted column held a row's only cell of its own beside a vertical merge, the direct deletion left a row of nothing but that merge: the saved package wrote it as a row of `w:vMerge` continuations, which reopens with the merge split apart, so the reopened snapshot, `getContent()` and `docxToMarkdown` disagreed on its spans. Such a row is now removed and the merge closes over one row fewer, as accepting the tracked deletion already did. Accepting a tracked deletion of a cell that spans several rows now removes the cell instead of moving it into the row below.

Other table operations now keep the saved grid coherent too:

- `mergeTableCells` over whole rows removes the rows left without a cell of their own, and the merged cell spans the rows that remain. A tracked vertical merge that would leave such a row is refused (`unsupportedBlock`).
- `deleteTableRow` through a cell wider than one column shortens a vertical merge further right in the row, which it left one row too long.
- `insertTableColumn` widens `w:tblGrid` with the new column, which kept the old grid.
- A vertical merge's continuation cells take the merge's current width when a column is inserted or deleted through it, instead of the width they were read with.
