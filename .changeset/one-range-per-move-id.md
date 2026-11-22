---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A document holds one range per move-range id. Inserted content leaves out a copy of a move range boundary that stands in the document, in every mode and on paste, as a copy of moved text is no part of the move; content pasted over its own source, or put back by an undo, keeps its boundaries. A part that repeats a range is repaired on open: its first range is kept, and each further one is dropped and reported as `dropped-move-range`.
