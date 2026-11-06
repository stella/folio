---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A text box keeps everything it holds. Content controls, custom XML, bookmarks, `w:altChunk` and other block markup in `w:txbxContent` used to be left out of the model and were dropped from the saved file when the paragraph hosting the box was edited, along with any text in them. Text in a control or custom XML now shows in the box, and an `w:altChunk` shows its marker.
