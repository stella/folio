---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A `newList` numbering request on `insertAfterBlock`, `insertBeforeBlock` or `setBlockParagraphProperties` now applies in the experimental canonical session. It mints a list instance per operation and kind, as the default engine does. The `insertParagraphs` and `formatParagraph` editor intents take the new instance and create it first. `NewListInstance` is exported from `@stll/docx-core/ops`.
