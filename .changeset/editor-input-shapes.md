---
"@stll/folio-core": patch
---

The editor DOM sets an image's `src` only for `data:image` and `blob:` sources; other sources stay on the node and round-trip through the clipboard under `data-folio-src`. Pasted list HTML is restructured in an inert document. Display-list font families are escaped as full CSS strings. The dropdown picker offers one choice per stored value. `applySuggestions` reports a suggestion overlapping one applied after it as stale. Replace All copies the document once per batch. A headless reviewer keeps its latest 100 document-operation undo entries.
