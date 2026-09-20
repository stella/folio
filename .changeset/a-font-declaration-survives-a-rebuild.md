---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Rebuild `word/fontTable.xml` from the model without losing what the source declared.

A repack copies the part across byte for byte, so the reader's gaps only showed on the paths that build a package from the model: a package folio authors, and a style set carried into a new document. `w:font` now goes through the shared child dispatcher, whose handler map the compiler makes total over the children `CT_Font` declares, so `w:notTrueType` lands in the ordered sink instead of on the floor and an attribute the font's record has no field for rides its remainder.

Four things the model held or dropped are now written back. `w:charset` keeps the character set it names as well as the one it numbers, and a bare `<w:charset/>` — the default code page — is no longer written as no `w:charset` at all. The four `w:embed*` faces are written from the model with their `w:fontKey` and `w:subsetted`, which nothing wrote before although the relationship id was parsed: a rebuilt part pointed at no embedded font.

`FontInfo.charset` becomes `FontCharset` and the four `embed*` fields become `EmbeddedFontRef`, so the key an embedded face cannot be decoded without travels with the relationship that names it. That retires the second font-table reader that existed only because the model dropped the key; one reader owns the part.
