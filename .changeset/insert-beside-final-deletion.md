---
"@stll/folio-core": patch
---

In the canonical session, a direct batch that deletes a paragraph with no paragraph after it (at the end of the body or a cell, or directly before a table) and inserts a paragraph beside it now applies both, as the default editor does, instead of refusing the insertion as overlapping.
