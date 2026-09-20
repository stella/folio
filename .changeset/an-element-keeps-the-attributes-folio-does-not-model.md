---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep the attributes an element carried that folio has no field for, through a save and through the editor.

Word writes a revision-session id on nearly every paragraph, run, row and section (`w:rsidR`, `w:rsidRPr`, `w:rsidDel`, `w:rsidP`, `w:rsidRDefault`, `w:rsidTr`, `w:rsidSect`). folio rebuilt each of those elements from the model alone, so opening a document and saving it rewrote the whole revision history.

`Paragraph`, `Run`, `TableRow` and `SectionProperties` gain `preservedAttributes`, an ordered list of resolved `{ namespace?, name, value }` records. The decision of what to keep is made on the resolved namespace URI and local name, so a source that binds a second prefix to the WordprocessingML namespace does not get a second copy of an attribute the parser already read; the writer is handed the modelled attributes it is about to emit and drops any remainder entry that would spell one of them again, so a duplicate attribute cannot reach the part. A namespace declaration is never in the remainder, and neither is an attribute whose namespace the rebuilt part cannot bind.

The remainder follows the record: a paragraph, row or section the editor creates from scratch has none, an authored one's survives `toProseDoc`/`fromProseDoc` unchanged, and when a command splits a record in two the half that comes first in document order keeps it. A run has no record in the editor — it is text plus marks — so a run's remainder survives a save and not the projection.
