---
"@stll/folio-core": minor
---

Read a numbering level's `w:pPr` with the reader the other three owners of the set already share.

A level kept a private copy that took an indent and a tab list and let the other thirty-one children `CT_PPrGeneral` declares fall off the end of the walk, while the level's writer was already the shared one: a `w:pStyle`, a `w:jc`, a `w:spacing` or a `w:keepNext` on a level was read as nothing and written back as nothing.

The reader could not be shared before because `paragraphParser` and `numberingParser` import each other, so it now lives in `docx/paragraphProperties.ts`, which imports neither and which both import.
