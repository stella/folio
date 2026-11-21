---
"@stll/folio-core": patch
---

On the canonical session, refuse a table placed right after a paragraph whose break is pending deletion (`pendingParagraphMarkDeletion`): accepting that deletion would run the paragraph's words into the table's first cell. A table beside a paragraph an earlier operation of the batch joined away goes beside the paragraph that now holds it.
