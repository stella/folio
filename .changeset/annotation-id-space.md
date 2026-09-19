---
"@stll/folio-core": patch
---

Mint revision ids out of the whole annotation space. `w:id` on a comment, a bookmark, a protected range and a tracked change comes from one counter, but the save-time deduplication reserved only revision ids and handed out the lowest free integer, so renumbering a duplicated `w:ins` in a commented document could land on a live comment's id. The pass now reserves every annotation id (never claiming one, since a comment id legitimately repeats), and both editor allocators seed above the maximum of the whole space rather than their own kind.
