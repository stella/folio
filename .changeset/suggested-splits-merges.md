---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Suggested batches can now carry `splitBlock` and `mergeBlockWithNext` in both engines. Each records tracked changes stamped with the suggestion: rejecting gives the source back, and accepting adopts them. A canonical public split now keeps the paragraph's identity on the first half, as the default engine's split does. In both engines, rejecting a split's inserted mark now gives the paragraph back under its own id.
