---
"@stll/folio-core": patch
---

Markdown export writes a heading or quote paragraph inside a table cell as its text, since a table cell holds inline content only and the heading or quote marker would otherwise show as literal text in the cell.
