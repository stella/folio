---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep the range markers that stand between two blocks. `w:body`, `w:tc`, a header and an SDT's content all admit `w:permStart`, `w:customXml*Range*` and a comment or move range beside their paragraphs, and every block container dropped them on save: a protected range lost its `w:permStart` and the saved file came back unprotected. The markers are now captured verbatim and replayed where they stood, the way an SDT's sibling markers already were. They do not yet survive the editor round trip, which needs a zero-width node rather than a block attribute.
