---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

An insertion of a paragraph right before the table that stops a pending join is refused with `pendingParagraphMarkDeletion`: once accepted, the paragraph would run together with the words before it or the table's first cell. Other insertions beside a pending join are unchanged.
