---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

In the experimental canonical session, suggested batches can now carry `splitBlock` and `mergeBlockWithNext`. Each records tracked changes stamped with the suggestion: rejecting gives the source back, and accepting adopts them. A canonical public split now keeps the paragraph's identity on the first half, as the default engine's split does. Rejecting a canonical split's inserted mark gives the paragraph back under its own id.
