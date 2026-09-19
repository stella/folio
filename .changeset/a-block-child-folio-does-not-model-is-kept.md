---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every block-level child folio does not model, where it stood, through the editor as well as through a save.

`w:body`, `w:hdr`, `w:ftr`, `w:tc`, an SDT's content and a note body share one walk, and it modelled paragraphs, tables and content controls and let the rest fall off the end. A `w:permStart` between two paragraphs is the whole of a document-protection range; `w:altChunk` is an entire imported document; `m:oMathPara` is a display equation. The walk now goes through the shared child dispatcher, whose handler map the compiler makes total over the children the schema declares for a block container and whose default is the verbatim sink.

`BlockContent` gains a `preservedBlock` member holding the captured markup, and the editor gains a zero-width `preservedBlock` node for it. Position is structural on both sides: the capture sits between the same two blocks in the model, in the ProseMirror document and in the saved part, so inserting, splitting or deleting a neighbour moves it the way a reader would expect and nothing has to keep an index honest.

`Paragraph`, `Table` and `BlockSdt` lose `rawMarkersBefore` / `rawMarkersAfter`, the narrower mechanism this replaces: it kept only sixteen range-marker names, dropped them when the container held no block at all, and had no editor leg, so a document that survived an untouched save lost the markup the moment anybody opened it. `Footnote.content`, `Endnote.content` and `TableCell.content` are now `BlockContent[]` rather than hand-written copies of it.
