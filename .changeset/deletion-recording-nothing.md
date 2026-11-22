---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Deleting a paragraph with nothing to change or record, as an empty one alone before a bookmark or one holding only range markers whose mark is already pending deletion, is a no-op or refused instead of reported as applied
