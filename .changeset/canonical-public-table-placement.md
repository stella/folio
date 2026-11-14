---
"@stll/folio-core": patch
---

The canonical session applies the public `insertSignatureTable` operation, directly or as a tracked change, and places a table inserted from inside a table beside the outermost one, against the paragraph on that side of it, as the default engine places it. A table inserted beside a paragraph whose break is pending deletion, where no pending join reads that side, goes beside the paragraph as it stands, as in the default engine.
