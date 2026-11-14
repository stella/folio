---
"@stll/folio-core": patch
---

The extension registry's `undo` and `redo` commands step a canonical session's own history, as the undo keys and the editor API do, so a toolbar or host that runs them by name undoes and redoes on the canonical engine.
