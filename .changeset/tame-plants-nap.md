---
"@stll/folio-core": patch
---

fix(core): apply table style paragraph spacing to cell paragraphs

Table cells now inherit paragraph spacing (space-after, line spacing,
contextual spacing) from the enclosing table style's `w:pPr` — and from the
applicable `w:tblStylePr` conditional region (first row, banding, etc.) —
instead of falling through to `docDefaults`. Per ECMA-376 §17.7.2 this table
style layer sits between docDefaults and the cell paragraph's own style
chain/direct formatting, so an explicit paragraph style or direct spacing on
the paragraph still wins. Previously, cell paragraphs with neither a
`w:pStyle` nor a direct `w:pPr` picked up the document's default paragraph
spacing (e.g. Word's default ~10pt space-after and 1.15x line spacing)
instead of the table style's typically compact spacing (e.g. `TableGrid`'s
zero space-after / single line), inflating table row heights.
