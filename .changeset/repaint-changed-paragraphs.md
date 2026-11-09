---
"@stll/folio-core": patch
---

Typing repaints only the paragraphs it changed on a page: the painter keeps every node on the page that shows what a fresh painting of it shows, taking new positions where an edit before it shifted them, instead of replacing the page's whole content area.
