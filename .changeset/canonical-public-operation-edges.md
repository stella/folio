---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

Canonical public operations: an inline insertion reads CRLF and CR as one line break; an insertion beside a paragraph in a table cell goes beside the table; an insertion beside a paragraph the same batch deletes applies when a block follows that paragraph; a move whose other half does not apply saves as a plain insertion or deletion, reported unpaired; an undone insertion's paragraph id is not reused; and an open header, footer or note editor keeps its caret through an operation on its story. A list reference that names no list stays out of lists started later.
