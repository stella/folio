---
"@stll/folio-core": patch
---

A merge whose next paragraphs are pending deletion with a table after them is refused in every mode, as a merge right before a table is, and blocks inserted on either side of a table pending deletion share one gap in the order the batch names them, so tracked edits accepted read as the same edits applied directly.
