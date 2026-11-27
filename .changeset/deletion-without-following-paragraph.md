---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

A suggested deletion of a paragraph that ends a section, followed by bookmark or comment range markers, is now recorded and resolves as the direct deletion does, its section break going with its mark. Deleting a paragraph right before a content control or kept block that ends its container now empties the paragraph in place, directly or as a suggestion, instead of being refused. Section property changes on a paragraph inside a content control or table cell now accept and reject. A direct paragraph insertion after a pending inserted mark that rejecting it could not reproduce is refused as a request to resolve that insertion first.
