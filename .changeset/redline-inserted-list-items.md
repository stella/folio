---
"@stll/folio-core": patch
---

`generateRedlineDocx` inserts a paragraph with the revised version's numbering, style and direct paragraph formatting instead of its anchor's, so accepting the redline keeps an inserted list item's bullet or number. Numbering the base package lacks is copied in, under a fresh id where the base uses that one for another list.
