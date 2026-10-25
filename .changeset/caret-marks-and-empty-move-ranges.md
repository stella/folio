---
"@stll/folio-core": patch
---

A document operation's replacement text no longer takes the caret's stored formatting when it replaces text away from the selection, such as an emptied heading's bold and size, and resolving a move that leaves its range empty keeps the range in the shape a saved package reopens with.
