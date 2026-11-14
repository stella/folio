---
"@stll/folio-core": patch
---

The canonical session refuses an editing operation naming a paragraph style or a numbering instance the document does not define, with the same `missingStyle` or `missingNumbering` skip as the default engine, instead of applying it.
