---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

Open every document that carries an explicit page-break run. `w:br w:type="page"` is an ordinary run child, so Word writes one in a bordered, framed or outlined paragraph, inside a table cell or a text box, and beside any inline kind. Folio refused several of those shapes at conversion and again at layout, which meant the document could not be opened in the editor, laid out or exported to PDF at all. They now project, save and round-trip; where layout can only approximate the break's owner, it says so through the parse-warning channel under the new `page-break-projection-approximated` code instead of throwing. `UnsupportedDocxToProseMirrorConversionError` goes with the last refusal that raised it.
