---
"@stll/folio-core": patch
---

Keep `w:cols/@w:sep` as the section authored it. The parser recorded only `true` and the serializer emitted only `true`, so an explicit `w:sep="0"` was read as an absence and written back as one. The attribute was also missing from `serializeColumns`' bail-out condition, so a `w:cols` whose only stated setting was the separator lost the whole element rather than the one attribute. `@w:equalWidth`, which already round-tripped correctly, joins it in the reserved-value registry so both column toggles are recorded against the reader that owns them.
