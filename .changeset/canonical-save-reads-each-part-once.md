---
"@stll/folio-core": patch
---

Saving with the canonical engine is faster: a save reads the generated document part once instead of twice, and reads the source part once per opened document instead of twice per save.
