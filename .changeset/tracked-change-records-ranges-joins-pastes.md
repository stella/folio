---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A suggested deletion now records comment ranges it crosses or holds whole, keeping a standing endpoint in place (inside a hyperlink too), and can delete an empty hyperlink or content control on its own. A suggested paragraph join records an inline merge depth, a suggested paste keeps the review and content-control records the pasted paragraphs carry, and a property change on a field's instruction run can be accepted or rejected. Removing a paragraph that owns an empty formatting field now undoes exactly, and a story-crossing range is refused as a cross-block range.
