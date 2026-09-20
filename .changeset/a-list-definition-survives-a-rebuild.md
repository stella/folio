---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Rebuild `word/numbering.xml` from the model without losing what the part defines.

A repack copies the part across and splices single definitions into it by id, so the reader's gaps only showed on the paths that build a package from the model. Measured over the cached public corpus, a rebuild used to lose 45 distinct element and attribute slots across the sampled packages that carry the part; it now loses none the part itself owns.

`w:numbering`, `w:abstractNum`, `w:lvl`, `w:num` and `w:lvlOverride` go through the shared child dispatcher, whose handler map the compiler makes total over the children each content model declares. A `w:numPicBullet` and the `w:numIdMacAtCleanup` high-water mark land in the ordered sink in source position, and an attribute a record has no field for — `w15:restartNumberingAfterBreak`, `w15:durableId` — rides its remainder.

`AbstractNumbering` gains `nsid` and `tmpl`, the identity Word recognises a list template by across documents; `ListLevel` gains `tplc`, `tentative`, `pStyle`, `lvlPicBulletId` and the `w:null` flag `w:lvlText` may carry. `lvlJc` widens to the whole `ST_Jc` enumeration `CT_Jc` declares, so a justification folio has no marker layout for is carried rather than taking the element with it; the three alignments layout does have are resolved from it. `NumberFormat` gains `bahtText` and `dollarText`, the two `ST_NumberFormat` members the model omitted.

Three values came back different from the way they were written. `w:legacy` reads its own `w:legacy` attribute rather than a `w:val` the type does not declare, so an explicit "off" is no longer written as "on"; `w:legacySpace` and `w:legacyIndent` spelled with a unit resolve to the twips they count; an explicit `<w:isLgl w:val="0"/>` stays off. An empty `<w:pPr/>` or `<w:rPr/>` is written as the empty element the source wrote, not as an absent one, and a `w:lvl` whose `w:ilvl` names no level is kept as the definition it is while resolving to no level.
