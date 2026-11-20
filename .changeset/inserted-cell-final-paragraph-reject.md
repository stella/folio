---
"@stll/docx-core": patch
---

Rejecting an inserted cell-final paragraph mark whose paragraph is left empty removes the paragraph and lets the paragraph before it end the cell, instead of leaving an extra empty paragraph in the cell.
