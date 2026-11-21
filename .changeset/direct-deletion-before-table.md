---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A direct deletion of the paragraphs before a table or block, with no table before them, removes them as accepting their tracked deletion does; the boundaries they hold go to the paragraph before, else into the block.
