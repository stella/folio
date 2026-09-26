---
"@stll/folio-cli": minor
---

Open and edit `.docx` files in VS Code: the Folio DOCX extension (whose version follows this package) now opens `.docx` files in folio's editor by default. Saves go through `folio save`, keep a backup of the previous version, and hold the editor lease while there are unsaved edits, so an agent's write saves them first instead of failing. The editor replaces the read-only preview: **Folio: Open Read-Only** opens a document with editing off. A `workbench.editorAssociations` entry that names the old `folio.docxPreview` view type no longer matches an editor; VS Code falls back to its default for `.docx`, which is now folio's editor, and the entry can be removed.
