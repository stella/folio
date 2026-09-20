---
"@stll/folio-core": minor
---

Carry a table's and a row's verbatim sink through the editor, not only through a save.

`Table.preserved` and `TableRow.preserved` hold the markup a `w:tbl` carried beside its rows and a `w:tr` beside its cells — a bookmark or permission boundary, a proofing error, a custom-XML revision range — with the count that places each back between the same two siblings. The save leg wrote them; `toProseDoc` had nowhere to put them, so opening a document and saving it dropped them. A `w:bookmarkEnd` written after a table's last row is the case that shows: losing it leaves the `w:bookmarkStart` in a cell with no end.

`TableAttrs` and `TableRowAttrs` gain `_preserved`, carried by reference the way `_preservedAttributes` is, so a table or row the editor created has no sink and a copy does not inherit one. `preservedSinkCarriers.ts` asks the question once per model record that declares a sink, over a union derived from the model rather than listed, so the next sink cannot reach the editor without an answer.
