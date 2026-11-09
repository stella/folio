---
"@stll/folio-core": patch
---

Layout looks for page breaks inside paragraphs and for numbered cross-reference fields only in the parts of the document that hold them, instead of walking every node on every key.
