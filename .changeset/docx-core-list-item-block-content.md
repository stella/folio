---
"@stll/docx-core": patch
---

`compileMarkdownToContent` no longer drops a table, code block, or blockquote nested inside a list item; it keeps the content as a following block and reports the flattening on the new `MarkdownContent.warnings` field.
