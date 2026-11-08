---
"@stll/folio-core": patch
---

The experimental canonical session now runs the editor's column and cell commands: insert a column left or right, delete columns, merge cells, split a cell and distribute columns. Each saves what the editor would save and undoes exactly. In suggesting mode, inserting a column, splitting a cell and distributing columns are tracked; deleting columns and merging cells refuse and change nothing.
