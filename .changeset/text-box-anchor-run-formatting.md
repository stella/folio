---
"@stll/folio-core": patch
---

A text box's run keeps its own formatting when the paragraph holding it is restyled, so a tracked merge or style change that names a new paragraph style no longer writes the old style's look into that run as untracked direct formatting, and rejecting it after a save restores the document.
