---
"@stll/folio-core": patch
---

`toMarkdown` keeps a comment whose start stands inside a revision wrapper, as it does one whose start stands beside it. Restoring a paragraph's properties, as a tracked merge's hand-over does, gives its body runs what a reopen reads from the paragraph mark, so those words read the same after a save.
