---
"@stll/folio-core": patch
---

A merge applied directly into a paragraph pending deletion right before a table keeps the merged paragraph apart from the table once accepted, and blocks inserted on either side of a table pending deletion share one gap in the order the batch names them, so tracked edits accepted read as the same edits applied directly.
