---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

Rejecting a paragraph-property change keeps every indentation the restored properties state, in both sessions, including a value equal to the one the paragraph's numbering supplied before. `registerStatedParagraphFormatting` and `paragraphStatesItsFormatting` mark and read formatting an edit gives as stated.
