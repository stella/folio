---
"@stll/folio-core": patch
---

Rejecting a list pasted while suggesting over words running to a paragraph's end restores that paragraph as it was: the paragraph that keeps the original break now records the properties the break had, so it no longer stays a list item.
