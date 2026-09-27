---
"@stll/folio-core": patch
---

Edits to a footnote or endnote that change its shape now reach the saved package. Deleting a note paragraph with no other edit, deleting a table, row or column, merging or splitting cells, and a tracked table deletion's row revisions used to report applied and then be dropped on save, because the note part was patched only at the paragraphs an edit changed. A note whose paragraphs were added, removed or replaced, or whose XML between paragraphs changed, is now rewritten whole from the model.

`FolioDocxReviewer.getComments()` reads comment anchors from every story, not only the body: a comment in a header, footer, footnote or endnote reports its anchored text and block id, and each thread gains a `story` handle naming the story that holds its anchor.
