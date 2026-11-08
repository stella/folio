---
"@stll/folio-core": patch
---

A paragraph mark's run properties (`w:pPr/w:rPr`) format the paragraph mark alone, as ECMA-376 specifies: no run of the paragraph shows them, and a save no longer writes properties turned off on runs for the mark's sake. Text typed into an empty paragraph takes the mark's formatting as its own direct formatting.
