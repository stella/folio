---
"@stll/folio-core": patch
---

Opening a document in a canonical session reads its package once to give paragraphs, content controls and revisions unique ids, and writes the package at most once, instead of reading and writing it once per kind of id.
