---
"@stll/folio-core": patch
---

The experimental canonical session now applies block-level public operations: `deleteBlock`, `insertAfterBlock`, `insertBeforeBlock` and `setBlockParagraphProperties`. They work in direct, tracked and suggested batches, and leave the document the default engine leaves, pending, accepted and rejected. Where the session cannot yet match that document, the operation refuses with a reason that names the case. This covers numbering, tracked moves, inserted page breaks, comments on these operations, multi-line or inline-break text, formatting scopes, insertion after a paragraph no paragraph follows, and deletion at the end of a block list or before a table.
