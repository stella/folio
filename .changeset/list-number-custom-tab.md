---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Send the tab after a list number to a custom tab stop between the number and `w:ind@left` instead of the hanging indent, apply the numbering level's `w:pPr/w:tabs` to its paragraphs, and honour `w:doNotUseIndentAsNumberingTabStop`.
