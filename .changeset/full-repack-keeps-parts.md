---
"@stll/folio-core": patch
---

A save now carries every part of the source package byte for byte — embeddings,
media, custom XML, fonts, macro projects, ActiveX controls, parts folio does not
model — and a macro-enabled document keeps its main-part content type. The only
entry a save refuses is one whose path would escape the package, and its
relationships and content-type entries leave with it, so a saved package never
references a part it no longer holds.
