---
"@stll/folio-core": patch
---

A canonical session reuses the editor nodes of every table an edit left, converting a table again only when it changed, the facts kept for its paragraphs changed, or a story fact it read (counters, open comments, bookmark and comment ids) changed.
