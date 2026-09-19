---
"@stll/docx-core": patch
---

Markdown headings compile to paragraphs carrying `w:outlineLvl` as well as a `HeadingN` style id, so the result is classified as a heading by outline level even when merged into a document whose own heading styles are localized and that id resolves to nothing.
