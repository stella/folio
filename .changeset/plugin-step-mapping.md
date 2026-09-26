---
"@stll/folio-core": patch
---

Large multi-step edits apply faster: a paste, replace-all, accept-all, AI edit batch or document compare no longer makes the paragraph change tracker, paragraph-id allocator, base-direction detection and run identity walk every paragraph through every step of the transaction.
