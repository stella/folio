---
"@stll/folio-core": patch
---

A `word/comments.xml` that defines two comments under one `w:id` now parses. The body addresses a comment by that id, so Word resolves every marker naming it to the first definition; folio keeps the first, drops the later ones no marker can address, and reports it as a parse warning. Footnotes and endnotes keep the first definition of a repeated id too, which their id index already did.
