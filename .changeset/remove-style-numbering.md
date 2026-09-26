---
"@stll/folio-core": patch
---

An operation that takes a paragraph out of a list its style numbers (`setBlockParagraphProperties` with `numbering: null` or `listLevel: null`, or an insert with `numbering: null` that keeps a numbered style) now states the cancellation (`w:numId="0"`), as the editor's list commands do. The numbering no longer comes back from the style when the document is saved.
