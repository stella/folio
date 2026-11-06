---
"@stll/folio-core": patch
---

A field whose result is in more than one format now behaves like one with a plain result. Deleting it with Track Changes on no longer stops the document saving, and replacing all of its result works. An edit that cuts into its result is refused rather than reporting applied and being lost on save.
