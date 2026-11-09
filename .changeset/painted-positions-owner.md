---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The editor positions a painted element carries are read and written through one module, `utils/paintedPositions`, and a lint rule keeps every other reader off the attributes. Painted output is unchanged, except that a table cell whose last block has no end position no longer carries the text "undefined" as its end.
