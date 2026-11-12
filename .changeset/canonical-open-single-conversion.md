---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

Opening a document in the canonical session converts it to ProseMirror once: the pages convert their own copy only while no editor view holds the document, as after a refused seed.
