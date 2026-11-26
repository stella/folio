---
"@stll/folio-core": patch
---

Opening a long document collects the fonts its first layout needs faster: each family is resolved once however many runs name it, and the characters to load are gathered in one pass over the text.
