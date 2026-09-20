---
"@stll/folio-core": patch
---

Keep an empty `w:rPr` on every owner of a run property set.

`w:rPr` is optional on a run, on the paragraph mark, on a style, on a numbering level and inside either kind of `w:rPrChange`, so a producer that wrote `<w:rPr/>` stated something an absent element does not. The one reader keyed the record on the properties the element yielded and answered `undefined` for one that yielded none, so the empty element reached no model and the one writer put nothing back. The carrier is now the element: the reader returns an empty record for an element that exists, and the writer writes the empty element exactly when the record is present.

The scan argues for it at every owner. Across the 5335 packages in the public corpus, `<w:rPr/>` appears 5121 times on a run, 3890 on the paragraph mark, 2419 on a style, 674 on a numbering level and 31 inside a `w:rPrChange`, and Word is among the producers of each. On the paragraph mark the case is sharper still: an empty one cannot state formatting, but it is where `w:rPrChange` and the mark's `w:ins`, `w:del`, `w:moveFrom` and `w:moveTo` live, so writing one is pure presence.

The paragraph mark's emission gains the third answer this needs: `undefined` for a mark that carried no property set, `""` for one that carried an empty one. `ParagraphFormatting.runProperties` is written only by the parser, so a present-and-empty record means the source had the element and nothing else can invent one.
