---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Rejecting the added break of an emptied paragraph that ends a block content control removes the paragraph whatever stands before it, and a paragraph whose break was rejected there while its inserted words stood is removed once those words are rejected too, so the inserted paragraphs' words and breaks resolve alike in any order; the marker is saved with the words.
