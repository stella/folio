---
"@stll/folio-core": patch
---

A canonical session holds each committed projection as the document its transaction ended on, so the next commit recognises the editor state as current without comparing every node.
