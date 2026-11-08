---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A complex field whose begin and end both sit inside a hyperlink is now one child of the link (`Hyperlink["children"]` admits `ComplexField`). A table of contents entry is one link over the heading text, a tab and a PAGEREF page number; the page number used to be read as loose field-character runs. The editor dropped those runs on save, so an edited entry lost its page number field, and the experimental canonical session refused the document. The field now saves inside the link it was read from, the editor round trip keeps it there, and the canonical session opens such documents. Applying a link over a selection that holds a field still leaves the field between the linked runs.
