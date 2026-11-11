---
"@stll/folio-core": patch
---

Layout blocks keep their ids across edits: a block kept or moved by an edit keeps its id, a part of the document that converts again takes back the ids it had, and new ids never repeat one still in use.
