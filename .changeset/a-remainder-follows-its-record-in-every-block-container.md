---
"@stll/folio-core": patch
---

Hold the attribute remainder to one record per block container, a text box included.

`keepOneAttributeRemainderPerRecord` gives an authored `w:rsid*` set to the first record holding it and to no other, so the half of a split paragraph the editor created claims no revision session. Its walk named paragraphs, tables and content controls and fell through on everything else, and `w:txbxContent` hangs off a shape inside a run rather than off a block child: no record inside a text box was ever reached. Splitting a paragraph or a row in one wrote the source paragraph's `w:rsidR` onto both halves.

The walk is now `visitBlockTreeRecords`, one traversal that enters a cell, an SDT's content and a shape's text body alike and is exhaustive over the block union, so a new block kind cannot be added without deciding what it holds. `visitDocxParagraphs` is the same traversal with its own pruning.
