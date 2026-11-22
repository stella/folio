---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Inserted content keeps one range per move-range id: a copy of a move range boundary that stands in the document is left out, in every mode and on paste, as a copy of moved text is no part of the move. Content pasted over its own source, or put back by an undo, keeps its boundaries.
