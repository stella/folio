---
"@stll/folio-core": minor
---

Keep an edit local in a package whose producer wrote no `w14:paraId`. The
selective save keyed paragraph identity on the id alone, and folio mints one for
every paragraph that arrives without one, so in a LibreOffice, Google Docs,
python-docx or docx4j document every paragraph looked new: the splice was
declined, the whole of `word/document.xml` was rebuilt, and the minted ids were
written to disk. `resolveParagraphIdentities` now decides each paragraph's
identity once — `authored` when the source part writes the id, `minted` when it
does not — and the patcher consumes that union exhaustively, addressing an
authored paragraph by id and a minted one by its ordinal, which the part's
remaining authored ids prove. A save no longer stamps a minted id into the
package; `ensureParaIds` remains the pass that gives a package ids, at ingest,
when a host asks for it.
