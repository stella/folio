---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

`insertTableRow` no longer drops a `cellTexts` value. Where a vertical merge crosses the insertion point, the merge grows through the new row and keeps its column, so the row has fewer cells than the table has columns; a value with no cell to go to was silently discarded from an operation reported as applied. Such an operation is now refused before anything is written, with the new skip reason and issue code `payloadDoesNotFit` (`retryable: false`, `recovery: "changeTarget"`), and `insertTableColumn` reports the same reason where a horizontal merge leaves the new column fewer cells than values. Skipped operations and issues carry an optional `message` naming the values that do not fit (`cellTexts[1] "Y"`).

A row inserted inside a vertical merge is now also accepted in tracked-changes mode, as it already was directly: the merge grows through the inserted row, accepting keeps it, and rejecting removes the row and shortens the merge again. A comparison's template row is no longer placed where a vertical merge crosses the insertion point, since its cells would overlap the merge.
