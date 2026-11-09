---
"@stll/folio-core": patch
---

Typing no longer walks every node of the document to check bookmark ranges before each layout; only the parts of the document that hold a bookmark are walked.
