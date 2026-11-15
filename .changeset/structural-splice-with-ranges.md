---
"@stll/folio-core": patch
---

A structural save (after Enter, a join or a paragraph deletion) now patches documents that contain bookmarks, fields, comments, permission ranges or tracked moves, instead of rewriting the whole package. It writes the whole package only when the edit would split one of those ranges or fields, and the measurement hook names that reason.
