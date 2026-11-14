---
"@stll/folio-agents": patch
"@stll/folio-core": patch
---

A public operation that edits the properties of a paragraph whose mark is pending deletion skips as `pendingParagraphMarkDeletion` in the canonical session too, as in the default engine, instead of the generic `pendingDeletion`. The skip is an ordinary rule of the public operations and no longer names a canonical gap.
