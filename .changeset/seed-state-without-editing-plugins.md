---
"@stll/folio-core": patch
---

The state the editor paints a document from before its canonical session opens carries no editing plugin of the extensions (ProseMirror history, the suggestion plugin, paragraph tracking, paragraph ids, table editing): it only paints until the session's state replaces it.
