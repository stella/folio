---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A suggested column deletion now keeps the table grid while it is pending and marks each deleted cell (`w:cellDel`), as the editor's own command does; accepting removes the column with its grid entry and rejecting restores the table exactly. The canonical session tracks a suggested column deletion instead of refusing it, and resolving revisions in a document that opens with a table no longer refuses.
