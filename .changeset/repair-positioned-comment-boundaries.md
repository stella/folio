---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Opening a package repairs a comment range boundary with no partner between a table's rows or a row's cells, where a reference anchors the comment, and a comment's repeated references (new parse warning `duplicate-comment-reference`); both sessions open such packages.
