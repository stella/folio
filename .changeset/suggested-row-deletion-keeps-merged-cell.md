---
"@stll/docx-core": patch
---

A suggested deletion of a table row that starts a vertical merge the next row continues leaves the merged cell's content unmarked, and accepting it hands the merge and its words to the next row, as deleting the row directly does. Rejecting a suggested row insertion that starts a merge hands the merge on the same way.
