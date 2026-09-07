---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A save that changed nothing writes a table's `w:tblPr`, `w:tblGrid`, `w:trPr` and `w:tcPr` back as they arrived, rather than rebuilding them from the typed model and dropping the conditional-format flags, the `w:tblGridChange`, and whatever else the model does not cover. The capture is re-parsed and checked against the model before it is used, so a `Document` edited in place is still honoured. `w:tcPr` and `w:tblPr` also stop gaining an inherited value — a border a table style supplied, a margin the table declared — as the cell's or table's own override, and an absent `w:hideMark` stops being written back as an explicit `w:val="off"`.
