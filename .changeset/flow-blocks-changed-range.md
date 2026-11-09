---
"@stll/folio-core": patch
---

Layout converts only the part of the document an edit changed: blocks before the edit are kept as they were, list numbering is followed only until it settles, and a paragraph the edit only moved takes the measurement key of the block it was moved from.
