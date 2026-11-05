---
"@stll/folio-core": patch
---

A tracked `mergeBlockWithNext` on a paragraph whose break is a move's destination retracts that break, as it does an inserted one, so rejecting every change no longer leaves an empty paragraph.
