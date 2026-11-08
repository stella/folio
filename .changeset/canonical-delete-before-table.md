---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A tracked `deleteBlock` of a paragraph that a table follows now applies in the experimental canonical session, as the default engine records it: its words and its own mark are marked deleted. `DELETE_BLOCKS` records this with a revision when the block list still ends with a paragraph. A new `deleteParagraphs` editor intent compiles to it.
