---
"@stll/folio-core": minor
"@stll/docx-core": minor
---

Dragging a selection that holds a table now moves the table, in Editing and Suggesting modes, as one undo step. A suggested move records the source as deleted and the copy as inserted. `@stll/docx-core` adds the `moveBlocks` editor intent, which moves blocks, tables among them, from a range to a target.
