---
"@stll/folio-core": patch
---

A tracked column insertion or deletion gives each cell a revision of its own, so `getChanges()` lists, and `acceptChange()`/`rejectChange()` resolve, one cell at a time, as after a save and reopen.
