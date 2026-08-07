---
"@stll/folio-core": minor
---

Extract DOCX tables as markdown rows instead of a flat row-major paragraph list, so a cell stays associated with its column. `ExtractedDocxParagraph` gains an optional `tableRow` describing which table a row belongs to and whether it carries cells.
