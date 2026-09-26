---
"@stll/folio-core": patch
---

A tracked deletion of text that is itself a pending tracked insertion (a reviewer or an agent revising a suggestion, e.g. two `replaceInBlock` operations on the same word in `"tracked-changes"` mode) now saves as a deletion inside the insertion (`w:ins > w:del`). The save used to write the insertion alone, so the deleted text came back on reopen (`"Signed in threefour copies."`).
