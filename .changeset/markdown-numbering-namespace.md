---
"@stll/folio-core": minor
---

`fromMarkdown` now synthesizes `document.package.numbering` for the ordered/bullet
lists it emits, so `createDocx(fromMarkdown(markdown))` no longer throws
`DocxModelValidationError: Numbering definition N is missing` and round-trips
through `docxToMarkdown` unchanged.

Added `mergeDocumentContent(target, source)`, a general helper for appending one
document's content onto another. It renumbers any `numId`/`abstractNumId` the
source carries to sit above the target's existing numbering range, so merging
`fromMarkdown`'s output into a styled preset (e.g.
`createStellaStyleDocumentPreset()`) can no longer collide with numbering the
preset already reserves — previously a markdown list could silently render with
the preset's own clause/definition numbering instead of a plain bullet/number.
