---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep the table properties a row overrides (`w:tblPrEx`), and the revision that records changing them.

Nothing read the element. The row's child walk called it `OWNED_ELSEWHERE` — "another reader owns this" — and no reader did, so the nine table properties a row may restate and the `w:tblPrExChange` beside them went on every save: 24 pairs in the survival census, from the row pair down through `CT_TblPrEx`, `CT_TblPrExBase` and `CT_TblPrExChange`. Word writes the element when a table is built by merging two, and a consumer reads it in place of the table's own properties for that row, so the loss restyled the row.

`TableRow.tablePropertyExceptions` holds it as the same `TableFormatting` the table carries, because `CT_TblPrEx` is the middle of `CT_TblPrBase` and the two are read by one set of handlers rather than two. `TableRow.tablePropertyExceptionChanges` holds the revision, as `Table.propertyChanges` holds `w:tblPrChange`. Both ride the ProseMirror row node, so the element survives the editor as well as a save.

`CT_Row` declares `w:tblPrEx` before `w:trPr`, and the serializer writes it there. The element is optional, so an empty one is kept rather than read as no exceptions at all: its presence is the value.
