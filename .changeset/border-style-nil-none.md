---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep `nil` and `none` distinct wherever a `CT_Border` is read or written. They are two members of `ST_Border`, not synonyms, and the build-from-scratch serializer rewrote `none` as `nil`. The four `parseBorderSpec` copies also collapse into one reader, so a border element with no `w:val`, an explicit `w:shadow="0"` and the page-border art relationship ids are now read the same way on the paragraph, style, table, cell and page tiers. The light grid a generated table gets when it declares no borders is unchanged, but is now named as the authoring default it is.
