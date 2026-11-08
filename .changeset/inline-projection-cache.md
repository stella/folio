---
"@stll/folio-core": patch
---

A canonical session reuses the editor-to-source position map of every paragraph an edit left, mapping again only the paragraphs whose records, editor node, or story facts they read (the comments open at their start, comment and bookmark ids) changed.
