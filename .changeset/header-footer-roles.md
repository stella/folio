---
"@stll/folio-core": patch
---

First-page and even-page headers and footers are parsed with their role (`hdrFtrType` `first` / `even`) from the section references that name them, instead of all as `default`. `getNotesAsText()` labels each header and footer with its roles, and, in a document with several sections, the sections that show it (`[header first (sections 1, 3)] …`).
