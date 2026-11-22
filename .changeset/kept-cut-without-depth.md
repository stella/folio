---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A kept range marker's cut (`KeptCut`, saved as `folio:keptCut`) no longer carries `depth`: no resolution read it. The field is gone from the model, the operation wire and the saved review history; a file saved with it opens as before, and its depth is read past.
