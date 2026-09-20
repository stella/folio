---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep the children a table holds beside its rows, between the same two rows, through a save.

`CT_Tbl` declares a permission range, a proofing error, the table-level comment and move ranges and the eight custom-XML revision ranges beside `w:tr`. The table walk read `w:tr` and unwrapped `w:sdt`; everything else it returned from. The walk now goes through the shared child dispatcher over a generated `table-content` set the compiler makes its handler map total over, with the verbatim sink as the default and `w:tblPr` / `w:tblGrid` marked as read elsewhere so neither is written twice. A `w:customXml` row wrapper is kept whole rather than dropped, and a `w:tbl` nested directly in a `w:tbl` reaches the sink through its default rather than being flattened into the rows around it.

`Table` gains `preserved`, the ordered sink whose `index` counts the rows that preceded a capture. A table-level child cannot be a row, so this is the sink case rather than the union case the inline levels use.

The editor leg stops at the save, as it does for the row sink one level down: the table node's children are rows, and a zero-width capture between two of them is not a row.
