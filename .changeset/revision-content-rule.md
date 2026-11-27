---
"@stll/folio-core": minor
"@stll/docx-core": minor
---

A suggested paste or move of a block content control is now refused under the document's `revisionContentShape` rule (reason `noRevisionRecord`), since a block-level content control takes no tracked change of its own, and the editor shows the new `editor.suggestionUnavailable` message in every catalog language; the same change made directly applies. A suggested deletion of paragraphs that would record nothing reports `noChange` instead of a capability gap.
