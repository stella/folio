---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Tracked changes no longer record the date and count of kept range markers (`keptFrom`, `keptCount` on resolution joins and deferred seam groups): resolution reads each marker's own cut. Review history version 2 never writes them; version 1 files are still read, and the markers they name are marked on open.
