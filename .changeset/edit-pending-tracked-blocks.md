---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

`replaceBlock` on a block pending deletion (its text and paragraph mark tracked as deleted, or its table row or cell) is now refused with the new `pendingDeletion` reason in every mode, instead of reporting the rewrite applied while accepting ran the new text into the next paragraph or dropped it with the row. `getChanges()` now lists the insertion of a saved `w:ins > w:del` (pending inserted text that a later tracked edit deletes) after a reopen, as it did before the save.
