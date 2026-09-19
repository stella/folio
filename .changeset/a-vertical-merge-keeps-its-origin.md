---
"@stll/folio-core": patch
---

Keep `w:vMerge="restart"` on a cell whose merge has no continuation to span.
The editor carries a merge origin as the cell's rowspan, which only exists once
a continuation joins it, so every column whose merge closed at a rowspan of one
lost its `w:vMerge` on save: a restart the table ends on, one a plain cell
interrupts, one another restart supersedes. A merged cell losing its origin
changes the table's visible structure.
