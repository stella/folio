---
"@stll/docx-core": patch
---

Rejecting a paragraph inserted before a block content control or custom XML that ends the body, holding its last paragraph, removes the paragraph whole instead of leaving an empty one, so a document saved with a pending suggestion there is the original.
