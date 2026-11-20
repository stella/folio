---
"@stll/docx-core": patch
---

A table whose rows end early is read with the empty columns they imply after them, so its revisions resolve; a column only such a row's `gridBefore` or `gridAfter` holds stays when the cell in it is removed, and the cell beside it takes the removed cell's column. Column edits still need every row to cover the grid.
