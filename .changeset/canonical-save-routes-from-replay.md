---
"@stll/folio-core": patch
---

Saving with the canonical engine is faster: the selective save places the changed paragraphs using the blocks the save already replayed, instead of scanning and parsing the whole rewritten document part again.
