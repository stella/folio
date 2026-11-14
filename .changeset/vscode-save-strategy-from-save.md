---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-cli": patch
---

Canonical views no longer carry the paragraph change tracker; a canonical save reads its changes from the session alone. `DocxEditorRef.save` takes `onDiagnostic`, which receives a canonical save's fidelity fallbacks instead of `onError`, and `canonicalSaveDiagnosticError` builds the error the adapters report for one. The VS Code extension's editor reports how each save wrote the package from the save itself: a declined selective patch is a full repack.
