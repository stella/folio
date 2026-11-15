---
"@stll/folio-core": patch
---

A document whose bookmark has a start without an end, or an end without a start, opens in the editor and in both reviewer engines, wherever the boundary stands: between blocks, in a paragraph, in a hyperlink, in a table cell, in a header or in a note. The boundary stays where the source put it through edits elsewhere, save and reopen. The editor's integrity pass still removes the other half of a pair an edit breaks, and a boundary an edit brings in alone.
