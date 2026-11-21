---
"@stll/docx-core": patch
---

A direct cell merge, cell split or column insertion or deletion applies on a table with pending review records and keeps them under their ids: a pending grid change takes the column too, and pending row and cell property changes take the spans and omissions the edit gives, so accepting or rejecting them gives the edit on the accepted or rejected table. The edit refuses only a record it cannot keep exact: a merge across rows of different pending insertions or deletions, a split of a span whose cell has pending property changes, a column edit beside a pending cell insertion or deletion that takes part of a column, or a pending change that records another span or merge than the edit rewrites. Tracked edits keep their refusals.
