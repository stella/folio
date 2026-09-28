---
"@stll/folio-core": patch
---

Deleting a section break by editing (Backspace or Delete across the paragraph that ends a section, a selection, cut or paste spanning it, or the remove-section-break command) now merges that section into the next one and saves: the joined paragraph keeps the break of the paragraph whose mark survived, as deleting the paragraph as a block does, and the change tracker records the removal the save accepts. In suggesting mode the deletion stays tracked and is recorded when accepted. Undoing it takes the removal back.
