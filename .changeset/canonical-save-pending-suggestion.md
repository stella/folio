---
"@stll/folio-core": patch
---

Saving in the experimental canonical session while a suggestion that adds paragraphs is pending no longer rewrites the whole document body: the paragraphs the suggestion adds are not in the file, so they no longer count as changes for the save to write.
