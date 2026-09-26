---
"@stll/folio-cli": minor
---

Add `folio save`, which commits a package a live editor serialized with the same lease, backup, and journal as a tool call. A write that finds an editor holding the lease with unsaved edits now asks it to save and release first, then applies on the saved version instead of failing with `locked`.
