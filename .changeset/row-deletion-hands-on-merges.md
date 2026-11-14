---
"@stll/docx-core": patch
---

Deleting a table row directly that starts a vertical merge the next row continues hands the merge to that row, with the removed cell's properties and content, instead of leaving a continuation without a start and dropping the merged cell's words. The merge ends there when no later row continues it.
