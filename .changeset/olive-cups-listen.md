---
"@stll/folio-core": patch
---

Keep body text out of the PDF outline. A paragraph whose `w:outlineLvl` is the reserved value 9, such as a `TOC Heading`, is no longer written as a bookmark nested nine levels deep.
