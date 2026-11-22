---
"@stll/docx-core": patch
---

Resolving tracked insertions that split one another gives the same records in any order. A pending insertion's cut is read between the records its own pieces stand between, past insertions made after it. Two pieces of an insertion meet once accepted or rejected, whatever stands beyond them, and as deep as their own records go. Range markers kept from a later insertion beside a cut stand in it with the cut's own; one kept from outside the record a later insertion was made in stays beside that insertion. An accepted deletion of one piece leaves a cut with the other piece.
