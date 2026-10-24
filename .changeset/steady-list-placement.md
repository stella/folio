---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

Inserted and merged paragraphs keep the list they ask for: a direct merge into a paragraph pending deletion, a tracked insertion before deleted closing paragraphs, and an insertion after a table in a batch that empties the closing paragraph now read as applying the operations one at a time or directly does. A merge that sets the merged paragraph's properties conflicts with another property edit of that paragraph in the same batch, a suggested operation whose id names a still pending suggestion gets a suggestion id of its own, and clearing a numbered paragraph style drops the style's numbering from the document model. A level-only override with no style numbering is no list in any reader, and clean Markdown reads a pending numbering removal as accepted.
