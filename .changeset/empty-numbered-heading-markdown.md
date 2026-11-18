---
"@stll/folio-core": patch
---

Write an empty numbered heading's visible label to Markdown (`# 1.`), as an empty list item keeps its marker; it was dropped while a whitespace-only one kept it.
