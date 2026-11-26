---
"@stll/folio-core": patch
---

`applyTemplatePlan` carries a source field that numbers itself (`AUTONUM`, `AUTONUMLGL`, `AUTONUMOUT`) as the field it is. Such a field holds no cached text, so the number it prints was missing from the output. A block reports where such fields stand in its text as `selfNumbers`.
