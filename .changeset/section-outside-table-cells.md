---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Document operations keep section breaks out of table cells: replacing or splitting a cell's paragraph, setting a table or its rows, inserting a row, creating a header or footer, or restoring story parts is refused with `sectionBoundary` when it would end a section in a cell, and so is setting a section endpoint on a cell's paragraph.
