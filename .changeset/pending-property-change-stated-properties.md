---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A tracked paragraph property change records, as the paragraph's current properties, what the paragraph states in its own properties, as a reopened file reads them. A direct property edit of a paragraph with such a change pending keeps the change recording them, and undoing that edit restores the change exactly.

List normalization restores derived formatting and pending paragraph changes atomically during undo.
