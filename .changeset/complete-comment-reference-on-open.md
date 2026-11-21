---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Opening a package gives a comment whose range has no `w:commentReference` its reference after the range end on both sessions (new parse warning `missing-comment-reference`). A save writes it only into a paragraph it rewrites anyway, so an unedited save keeps the document part byte-identical.
