---
"@stll/folio-core": patch
---

Accepting or rejecting every change resolves the mark of a paragraph that holds a change a content control's lock keeps pending: the paragraph runs into the next one with the held change unchanged. Only a held proposal, which is restaged on the paragraphs it was made on, still keeps its paragraph's mark pending.
