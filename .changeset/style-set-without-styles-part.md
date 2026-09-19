---
"@stll/folio-core": patch
---

A DOCX with no `word/styles.xml` no longer fails style-set extraction. The package is valid and Word opens it on its built-in defaults, so the extracted set carries a minted default paragraph style. Where the source declared no `w:docDefaults`, that style carries the built-in Normal formatting too, because a set that names a default paragraph style is no longer a package a consumer applies its own built-in to.
