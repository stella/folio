---
"@stll/folio-core": patch
---

Carry a table's own properties through `compareDocx`: an inserted table brings the target's `w:tblPr`, `w:tblGrid` widths, `w:trPr` and `w:tcPr` (spans, merges, shading, borders, margins, alignment) and its nested tables; a removed table keeps them under the deletion marks, with the runs inside a deleted row marked too; a paired table, row or cell whose properties changed records `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange`. The round-trip self-check now compares the table model as well as the blocks, with a `table-geometry` cause. `w:tblPr` children are written in the order `CT_TblPrBase` declares, so a table carrying both `w:tblLayout` and `w:tblCellMar` validates.
