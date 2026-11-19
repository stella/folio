---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep a comment whose range closes between a table's rows or a row's cells. Editing or deleting a cell paragraph in such a table no longer drops the comment's reference, adds a second range end, or leaves a deletion that cannot be accepted or rejected.
