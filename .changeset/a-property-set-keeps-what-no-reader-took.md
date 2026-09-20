---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every `w:tblPr` and `w:sectPr` child folio does not turn into a typed value, in the order the content model declares.

Both property sets were walked by a reader per property with no branch for the rest, so a child whose value the reader did not admit and a child folio models nothing for fell off the end: 19 table pairs and 16 section pairs in the survival census. The section serializer had a second answer to the same question, returning the empty string when the element carried any unread child, which failed the whole save rather than one property.

Both now go through the shared child dispatcher with a handler map the compiler makes total over the schema's declared children, and a handler answers with what it took — a property the reader turned into nothing keeps its bytes. `TableFormatting.preserved` and `SectionProperties.preserved` hold them. `CT_TblPr` and `CT_SectPr` are sequences, so the sink records the schema ordinal rather than a count of modelled siblings, and both serializers merge modelled and captured children by it, reading the order from the generated declared-child list instead of restating it.

`w:tblCaption`, `w:tblDescription`, `w:tblStyleRowBandSize` and `w:tblStyleColBandSize` are authored values rather than markup nobody reads, so `TableFormatting` models them as `caption`, `description`, `rowBandSize` and `columnBandSize`.

A tracked property change no longer needs a non-empty snapshot to survive: `w:tblPrChange`, `w:trPrChange` and `w:tcPrChange` on a property set that states nothing of its own kept the author, the date and the id, and were dropped anyway.
