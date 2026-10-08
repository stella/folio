---
"@stll/docx-core": patch
---

Join paragraphs across tracked moved paragraph marks: a `w:moveFrom` mark removes the paragraph break in the current view like a `w:del` mark, and a `w:moveTo` mark removes it in the original view like a `w:ins` mark. Materialize `w:endnoteReference` as U+0002 in projected text, like `w:footnoteReference`.
