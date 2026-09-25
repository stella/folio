---
"@stll/folio-core": patch
---

Apply the default table style when `w:tblStyle` names a style the document does not define, and inherit table style borders (`w:tblBorders`, `w:tcBorders`) and cell margins (`w:tblCellMar`, `w:tcMar`) through `w:basedOn` side by side instead of replacing the whole set.
