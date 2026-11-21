---
"@stll/docx-core": patch
---

Rejecting the added break of an emptied paragraph that ends a block content control removes the paragraph whatever stands before it, so rejecting inserted paragraphs one at a time restores the original in any order.
