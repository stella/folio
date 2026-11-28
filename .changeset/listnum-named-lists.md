---
"@stll/folio-core": patch
---

A template plan keeps a LISTNUM field of the source as the field it is, as it keeps an AUTONUM: it holds no result to write as text, so its number was lost. A field that counts in a list by name takes that list's definition along into the output's numbering, where the template does not define a list of that name, without what ties it to its source package (its id, template code and style links).
