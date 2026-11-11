---
"@stll/folio-core": patch
---

Deleting a table right after paragraphs whose breaks are pending deletion is refused with `pendingParagraphMarkDeletion`: once accepted, the words before the table would otherwise run into the paragraph after it. Resolve the pending paragraph mark deletion first.
