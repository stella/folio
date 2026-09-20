---
"@stll/folio-core": minor
---

Carry `w:numPr` as a union through the editor attr

The persisted `numPr` and `numPrFromStyle` paragraph attrs hold
`ParagraphNumberingOverride`, the shape the model already carries, minted by
one codec. `toProseDoc` and `fromProseDoc` stop converting between two slot
pairs and a union at the editor boundary, and a recorded `w:pPrChange` stores
the same union with `null` still meaning "carried no numbering".

`FOLIO_YJS_ATTR_SCHEMA_VERSION` is 5. A stored version-4 snapshot is carried
forward by `migrateFolioYjsSnapshot` and by every load path: `w:numId` 0
becomes the cancellation, an `w:ilvl` without an id becomes the level-only
arm, the pair becomes a reference, and an element that stated neither slot
stays absent.
