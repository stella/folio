---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Setting, removing or inserting a hyperlink now works over a run with a pending formatting change, and over tracked insertions, deletions and moves. Before, the experimental canonical session refused these edits. A record the link edit cuts keeps its id on the first half, and the halves after it take fresh ids, as any cut does. Text inside a tracked change is linked inside that change, which is how the file nests it.
