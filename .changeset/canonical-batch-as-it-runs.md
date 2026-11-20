---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

Run a canonical batch's deletions and paragraph edits against the paragraphs as the batch has left them: a run deleted back to its container's start keeps its first paragraph, emptied; a property edit runs before the deletion that retires its paragraph's break; an edit of a paragraph an earlier operation joined away edits the one that holds it.
