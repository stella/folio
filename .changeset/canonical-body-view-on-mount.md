---
"@stll/folio-core": patch
---

In the experimental canonical session the editor creates its body view when the host syncs, instead of when a header, footer or note editor first reads a story. Reading a story's projection no longer creates the view.
