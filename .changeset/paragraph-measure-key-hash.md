---
"@stll/folio-core": patch
---

The paragraph measure cache keeps a 128-bit hash of what a paragraph measures instead of its text, so a long document's cache retains a fraction of the memory.
