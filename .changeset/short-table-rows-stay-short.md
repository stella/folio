---
"@stll/folio-core": patch
---

A table row that stops short of the table's width now stays as it is in the file. An edit anywhere in such a table added empty cells to the row, and could move the last row's text a column to the right, without a tracked change and without rejecting bringing it back; inserting a row could throw. Row and column edits now treat these positions like a declared `w:gridAfter`: a row left with no cells of its own is removed, including when a tracked column deletion is accepted. The editor's row, column, merge and split commands decline on such a table, as they do on one with a declared `w:gridAfter`.
