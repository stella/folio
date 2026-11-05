---
"@stll/folio-core": patch
---

`toMarkdown` and `docxToMarkdown` keep emphasis together across content that writes nothing, such as a stripped comment reference, instead of writing `**a****b**`, which reads back as literal asterisks.
