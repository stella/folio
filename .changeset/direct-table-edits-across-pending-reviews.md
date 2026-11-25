---
"@stll/docx-core": minor
---

Direct cell merges, cell splits and column edits apply beside more pending table reviews. A merge across rows with different pending insertions or deletions keeps each row's change on the words it moves; splitting a cell with a pending property change gives each new cell a record of its own; a column edit beside a partial cell insertion or deletion, or a pending span change, lands in the same column whether the reviews are accepted or rejected. Removing a row that alone continues a vertical merge ends that merge, and a resolved row's merged cell keeps its pending paragraph marks when it is handed on.
