---
"@stll/folio-core": patch
---

The canonical session's block cache checks the comments open around a kept block without copying them, and leaves them as they are past a block that opens or closes none.
