---
"@stll/folio-core": patch
---

A direct batch that merges a block into one it deletes joins across the deleted block, as applying the operations one at a time does, and refuses the merge where the deletions run to the story's end; tracked, such a merge no longer leaves its separator dangling at the end.
