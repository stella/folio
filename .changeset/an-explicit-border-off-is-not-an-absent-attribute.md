---
"@stll/folio-core": patch
---

Write `CT_Border/@w:shadow` and `@w:frame` when the border authored them, whichever state they authored. The serializer emitted the attribute only when the model held `true`, so an explicit `w:shadow="0"` came back from a save as an absence on every border position (paragraph, style, table, cell and page). Neither attribute carries an XSD default, so the two are not interchangeable. `parseOnOffAttribute` already kept all three states; only the emit collapsed them. The value written is now `1`/`0`, matching Word and the table and section serializers, rather than `true`.
