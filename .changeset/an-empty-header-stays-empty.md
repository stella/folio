---
"@stll/folio-core": patch
---

Keep an empty header or footer empty. The rebuild path added `<w:p><w:pPr/></w:p>` whenever the part came out with no blocks, on the premise that OOXML requires one: `CT_HdrFtr` holds a single `EG_BlockLevelElts` occurrence whose choice members are all optional, so a part with no block children is valid and is what Word writes for a blank header. Verbatim replay returned such a part unchanged, so the invented line appeared only after the document had been edited.
