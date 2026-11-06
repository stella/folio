---
"@stll/folio-core": patch
---

The canonical session no longer refuses a composition inside a tracked run because of how the browser's DOM reads that run back. The session keeps its own marks on commit, so only the composed text changes.
