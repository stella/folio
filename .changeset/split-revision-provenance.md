---
"@stll/folio-core": patch
---

A tracked change that holds a hyperlink now saves with its review provenance. OOXML nests the revision inside the link, so such a change is written as one element per stretch of text and link. Before this, the save refused such a change, and the experimental canonical session refused suggesting-mode edits that would create one, such as typing, Enter, deleting and cutting next to a link. Each piece of the change now carries a `folio:splitResolutionJoins` attribute in the ignorable folio review-history namespace, and the first piece carries the whole change's resolution joins. On reopen, the pieces are read back as the one change, under its first id and with its joins. A transparent wrapper split around the link gets the same treatment. A reader that does not know the attribute keeps every revision and drops only the joins. Unedited documents save byte-for-byte as before.
