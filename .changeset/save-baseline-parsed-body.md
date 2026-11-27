---
"@stll/folio-core": patch
---

An editor that opens a document itself compares its saves against the parsed body rather than a copy of it, so opening a long document keeps one copy of the body instead of two.
