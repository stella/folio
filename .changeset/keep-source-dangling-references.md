---
"@stll/folio-core": patch
---

A save no longer rewrites relationships, `r:embed` attributes or content-type overrides that named nothing in the source package. Only a reference the save itself broke is repaired. A relationship to a part whose entry name keeps its URI escapes is no longer mistaken for a dangling one.
