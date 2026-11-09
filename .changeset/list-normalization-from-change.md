---
"@stll/folio-core": patch
---

A list normalization in the canonical session starts at the first changed block, with the counters the previous normalization had there, and continues after the last changed block only while the counters before a block differ from those it had before. A block it normalized is kept as the block it became, so the next normalization reuses it.
