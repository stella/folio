---
"@stll/folio-agents": patch
"@stll/folio-core": patch
---

A `setBlockParagraphProperties` on a paragraph whose mark is pending deletion skips as `pendingParagraphMarkDeletion` in both engines as an ordinary rule of the public operations, and no longer names a canonical gap. A `replaceBlock` there follows the same rule in both engines: it skips as `pendingDeletion` only when the paragraph goes whole once its pending deletions are accepted.
