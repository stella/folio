---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Generate `NumberFormat` from `ST_NumberFormat`. It omitted `bahtText`,
`dollarText` and `custom`, and carried three `decimalZero{3,4,5}` members the
format does not declare: the parser minted them from a custom format's pad
width and the serializer wrote them back as a `w:val` no consumer can read.

A custom format is now held as `custom` plus the `@w:format` it counts by
(`ListLevel.numFmtFormat`), and written back as both. The three synthetic
values move to `CounterFormat`, the render vocabulary `ListRendering.numFmt`
and the editor's list attributes carry, which is never serialized.
