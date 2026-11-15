---
"@stll/folio-core": patch
---

On the canonical session, a paragraph whose mark is pending deletion takes the edits the default engine takes there: replacements inside it, formatting, a whole-text replacement and blocks beside it, which go beside the paragraph itself where no pending join reads it. It refuses what the default engine refuses, with the same reason: property edits (`pendingParagraphMarkDeletion`), a replacement of a blank paragraph that goes whole or of one in a deleted row or cell (`pendingDeletion`), and a block a pending join would run into a table after it (`pendingParagraphMarkDeletion`).
