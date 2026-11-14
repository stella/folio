---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Setting, removing and inserting a hyperlink, and generating a table of contents, are recorded as suggestions in the canonical session. A link change strikes each stretch of the changed text and inserts its changed copy after it, so rejecting brings the text back as it was and accepting leaves what the edit made directly leaves; a table of contents records each paragraph it inserts and the split it makes. A suggested link change that would leave a bookmark or comment marker alone inside a piece of a link, or copy a drawing, shape or note reference, is refused with a typed reason.
