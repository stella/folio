---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Model `wp:wrapPolygon` on `ImageWrap` and write the authored outline back, instead of a constant 21600-unit rectangle for every tight and through wrap. Wrap insets now record whether `wp:inline`/`wp:anchor` or the `wp:wrap*` child stated them, so a rebuild writes each one where it was authored.
