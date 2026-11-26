---
"@stll/folio-core": patch
---

`analyzeDocxBlocks` renders a list level in the `ordinal` number format by the language of the paragraph mark, as the reference consumer prints it: Czech and German write the number with a period (`1.`), French `1er` and `2e`, English `1st`. The language is the mark's own, its style's, or the document default. A label in a language whose ordinals are known can be written as text by a template plan; in any other language it stays as before and is not written.
