---
"@stll/folio-core": patch
---

The experimental canonical session now opens a document with a table row that starts late or ends early (`w:gridBefore` or `w:gridAfter`). Before, the whole document was refused. The cells the row does hold take typing and save as before. The space the row leaves out takes no input.
