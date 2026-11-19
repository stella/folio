---
"@stll/docx-core": patch
---

Resolve a tracked column insertion or deletion that only some rows take: a grid column goes only when every row loses its cell, and a row that loses a cell in a kept column widens the cell before it (or the one after it, at the start of the row), instead of refusing the review or the save.
