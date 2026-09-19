---
"@stll/folio-core": patch
---

Read a table width spelled as a percentage as one. `w:type` is optional on `CT_TblWidth` and the schema gives it no default, so folio read a missing one as `dxa` and `<w:tblW w:w="50%"/>` became 50 twips, then saved that way: a table half the page wide came back a hairline. `w:w` is `ST_MeasurementOrPercent`, so a `%` spelling is a percentage whatever `w:type` says; which slots admit it and what unit their number counts in now comes from the generated slot table the verbatim capture already uses, so the reader and the capture cannot disagree about one width.
