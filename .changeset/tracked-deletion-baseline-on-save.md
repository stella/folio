---
"@stll/folio-core": patch
---

Saving a package whose tracked deletions sit in paragraphs with their own paragraph-mark formatting no longer writes the language they inherit as direct run formatting, and an inserted paragraph no longer copies the source attributes (revision session ids) of the paragraph it follows.
