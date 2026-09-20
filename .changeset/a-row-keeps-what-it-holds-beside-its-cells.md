---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep the children a table row holds beside its cells, between the same two cells, through a save.

`CT_Row` declares a permission range, a proofing error, the row-level comment and move ranges and the eight custom-XML revision ranges beside `w:tc`. The row walk read `w:tc`, unwrapped `w:sdt` and carried a bookmark boundary into a neighbouring cell's paragraph; everything else it returned from. The walk now goes through the shared child dispatcher over a generated `row-content` set the compiler makes its handler map total over, with the verbatim sink as the default and `w:trPr` / `w:tblPrEx` marked as read elsewhere so neither is written twice.

`TableRow` gains `preserved`, the ordered sink whose `index` counts the cells that preceded a capture. A row-level child cannot be a cell, so this is the sink case rather than the union case the inline levels use.

The editor leg stops at the save: the table schema has no row-level node a zero-width capture could be, and an index recorded on the row node would drift the first time a column moved. The contract records those pairs as `editorProjection` rather than `neverParsed`, which is the difference between markup the model holds and markup folio never read.
