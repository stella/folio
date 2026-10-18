---
"@stll/folio-core": patch
---

Refuse a batch operation when an earlier accepted operation removes its target block, including merges and table removals, while preserving tracked paragraph-mark carriers as valid targets.
