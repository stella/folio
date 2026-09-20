---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Model `w:numPr` as a union

`ParagraphFormatting.numPr` and `numPrFromStyle` carry
`ParagraphNumberingOverride` instead of two optional slots, so the reserved
`w:numId w:val="0"` has no representation past the parse boundary and a level
stated without an id is a named arm rather than a half-filled pair. The
cascade fold `mergeParagraphNumbering` replaces the object spreads that used
to restate ECMA-376 17.3.1.19 at each tier.
