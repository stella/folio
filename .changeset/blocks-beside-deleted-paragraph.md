---
"@stll/folio-core": patch
---

Blocks a canonical batch puts before a paragraph it deletes, when a table or another block follows that paragraph, land in the requested order applied directly, as accepting the tracked batch leaves them. A deleted paragraph that would then stand between a table put before it and a table after it stays, emptied, in every mode, so the tables stay apart.
