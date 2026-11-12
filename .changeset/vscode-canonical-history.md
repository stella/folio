---
"@stll/folio-cli": patch
---

The VS Code extension's editor runs the canonical engine. Its undo and redo walk the engine's own undo steps, and the document's dirty state follows them.
