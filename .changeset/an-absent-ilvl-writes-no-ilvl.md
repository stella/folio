---
"@stll/docx-core": patch
---

Stop writing `<w:ilvl w:val="undefined"/>` for a paragraph that stated no level.

`serializeDocumentToDocx` wrote `w:ilvl` unconditionally beside `w:numId`, so a `numPr` carrying only an id produced an attribute value `CT_DecimalNumber` does not accept. An absent `w:ilvl` is level zero and is not the same bytes as a stated `w:val="0"`, so it stays absent.
