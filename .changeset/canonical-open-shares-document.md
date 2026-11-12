---
"@stll/folio-core": patch
---

Opening a document in the canonical session no longer copies it: the session shares the read-only document it opens, and reads the package's paragraph ids once.
