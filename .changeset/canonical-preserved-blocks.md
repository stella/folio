---
"@stll/folio-core": patch
---

The experimental canonical session now opens a document holding a block it keeps as opaque markup, such as an imported HTML chunk (`w:altChunk`). Before, the whole document was refused. The block is saved as it was, and a deletion across it removes the text on each side and keeps the block.
