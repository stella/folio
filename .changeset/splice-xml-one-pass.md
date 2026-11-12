---
"@stll/folio-core": patch
---

Opening and saving a large document is faster: every edit to a document part is applied in one pass over the part instead of rebuilding it once per edit.
