---
"@stll/folio-cli": minor
---

Open and edit `.docx` files in VS Code: the Folio DOCX extension (whose version follows this package) now opens `.docx` files in folio's editor by default. Saves go through `folio save`, keep a backup of the previous version, and hold the editor lease while there are unsaved edits, so an agent's write saves them first instead of failing. The read-only preview stays one click away.
