---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every declared child of a `w:trPr` and a `w:tcPr`, in the order the schema declares, through the editor as well as through a save.

Both property sets were read by name: ten of the row's fifteen declared children had an `if` and the rest had nothing, so `w:cnfStyle`, `w:divId` and `w:tblCellSpacing` went on every save, as did `w:hMerge`, `w:headers` and the structural revision a `w:tcPrChange` snapshot records. So did any child whose value the reader refuses — a `w:trHeight` of zero, a `w:vAlign` the enumeration does not admit, an explicit off. Both sets now go through the shared child dispatcher, whose handler map the compiler makes total over the children the schema declares, and `TableRowFormatting.preserved` / `TableCellFormatting.preserved` hold what no reader took a typed value from.

Each set is written by one call through `serializeSequenceChildren`, so the order is the generated declared-child list rather than the order of the serializer's statements: a `w:tcPr` whose children arrive out of `CT_TcPrBase`'s sequence comes back in it, and the sink's captures land between the same neighbours they were read between.

An empty `<w:trPr/>` or `<w:tcPr/>` is kept. Both elements are optional, so a producer that wrote one stated something an absent element does not, and a parser that keyed the record on the properties the element yielded deleted it on save.

`TableCellPropertyChange` gains `previousStructuralChange`: `CT_TcPrInner` declares the cell's insertion, deletion and merge, so a `w:tcPrChange` may record the cell as having stood inserted before the change, which is not the cell's current revision.
