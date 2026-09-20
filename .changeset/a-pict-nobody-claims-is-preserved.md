---
"@stll/folio-core": patch
---

Keep a `w:pict` whose VML shape looks like a watermark but belongs to nobody.

VML has no serializer, so an unresolved `w:pict` is either replayed from the bytes the parse captured or lost. `shouldPreserveRawVmlPict` declined for any shape `isWatermarkShape` recognised, on the premise that `watermarkParser` had claimed it. That premise was a guess about another module. The watermark reader claims a direct `v:shape` child of a `w:pict`, carrying a non-empty `v:textpath` or a `v:imagedata`, alone in its paragraph, in a header; a `v:oval`, a shape nested in a `v:group`, a shape sharing its paragraph with text and every `w:pict` in a footer or in the body all fall outside it. Those were declined by one owner and claimed by no other, so the artwork was dropped and the relationship its `v:imagedata` named stopped resolving. Verbatim part replay hid it until the part was rebuilt.

The decline is gone. The watermark owner removes its own artwork from the model, by emptying the paragraph it detached, so the run parser does not have to guess who claimed what; a watermark the header reader does claim is still written once.
