---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A direct `mergeBlockWithNext`, and a direct `deleteBlock` of a container's final paragraph that a paragraph precedes, leave the paragraph whose mark remains, the later one, as accepting the same edit tracked does: both engines keep its block id. The applied operation reports the join as `joined`, with the block that holds the joined text and the id it retired, and the operation's receipt names the first as updated and the second as deleted.
