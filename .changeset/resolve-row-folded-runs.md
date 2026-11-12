---
"@stll/folio-core": patch
---

The change list now reaches every revision a table holds. Resolving a tracked row by the id the list shows also resolves the run revisions it folds and the revisions in its vertically merged continuation cells. A merged cell's entry carries the revisions in the cells it folds, and a table grid revision is part of the table's property change, or lists as one when the table has none. Rejecting a deleted row, for example, restores all of its text.
