---
"@stll/folio-react": patch
---

`getDocument`, save and comment replacement in the default session copy the document once: the committed document is put together with the live body, numbering and notes first, and that whole is copied, instead of copying the committed document and the live stories separately.
