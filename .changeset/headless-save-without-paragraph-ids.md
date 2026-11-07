---
"@stll/folio-core": patch
---

`FolioDocxReviewer.toBuffer()` now saves a document whose paragraphs carry no paragraph identities when its default paragraph properties say "no numbering". It used to fail even without an edit.
