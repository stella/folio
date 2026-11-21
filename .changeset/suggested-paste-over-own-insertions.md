---
"@stll/folio-core": patch
---

Input while suggesting over pending revisions now reviews as the direct edit: a paste from inside one inserted paragraph into the next keeps the first paragraph's properties and the end paragraph's pending break, a paste over everything removes the table rows the author inserted, and Enter in a paragraph whose break is pending moves that break to the new paragraph. Withdrawing an inserted row shortens the vertical merges crossing it, as rejecting it does, and a table pasted over everything goes before a table the document starts with, every row of which is struck or withdrawn.
