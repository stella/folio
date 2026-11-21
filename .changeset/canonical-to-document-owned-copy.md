---
"@stll/folio-core": patch
---

On the canonical session, `FolioDocxReviewer.toDocument()` returns a document the caller owns: editing it leaves the session's later operations, saves and reads unchanged.
