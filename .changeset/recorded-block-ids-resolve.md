---
"@stll/folio-core": patch
---

Export `currentFolioBlockId`: the id a block recorded under an earlier version answers to now. A paragraph id above the 31-bit bound is brought into range on parse, so a block id recorded from such a paragraph before that stopped resolving; the mapping is a pure function of the id, so a host resolves a recorded id without the document.
