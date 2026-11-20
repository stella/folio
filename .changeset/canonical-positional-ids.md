---
"@stll/folio-core": patch
---

The canonical reviewer reads the paragraph ids it mints on opening as positional, so `compareDocxVersions` (now taking `session`) pairs id-less documents by content on either engine.
