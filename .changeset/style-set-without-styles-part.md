---
"@stll/folio-core": patch
---

A DOCX with no `word/styles.xml` no longer fails style-set extraction. The package is valid and Word opens it on its built-in defaults, so the extracted set carries one empty default paragraph style, which resolves through the consumer's built-in Normal exactly as the source's paragraphs did.
