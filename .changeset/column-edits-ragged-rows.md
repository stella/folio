---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Column edits work on tables whose rows end early or hold no cells: a row that stops short omits the columns after its cells, as a stated `gridAfter` does. Deleting a column takes a row with it when the row loses its only cell (directly, or when the suggested deletion is accepted), and removing the column that holds every cell removes the table. Rejecting a suggested column insertion on such a table restores it.
