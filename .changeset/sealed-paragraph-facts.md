---
"@stll/folio-core": patch
---

Typing in the canonical session no longer checks the kept facts of every paragraph before reusing its converted block or list normalization. A paragraph's kept facts are fixed once a cache keeps it, and an edit that changes them gets a new paragraph instead.
