---
"@stll/folio-core": patch
---

A run whose character style sets a toggle property (bold, italic, caps and the like) that the paragraph mark's own formatting also sets no longer loses it on save. Any edit in such a paragraph, and accepting or rejecting a change in it, wrote the property as turned off on every such run, so a bold heading came back plain.
