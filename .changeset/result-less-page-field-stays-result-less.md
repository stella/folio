---
"@stll/folio-core": patch
---

Keep a result-less `PAGE` or `NUMPAGES` field result-less on save. The save path wrote a literal `1` into the result of a field the author left empty, so a document reopened from folio said "1" where the source said nothing, whatever page the field sits on. Layout computes the number from the page it paints, so the invented result added nothing and changed what the document says. `proseDocToBlocks` no longer takes an `emptyFieldResult` mode and `EmptyFieldResultMode` is no longer exported: there is one behaviour now.
