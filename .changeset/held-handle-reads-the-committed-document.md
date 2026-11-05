---
"@stll/folio-react": patch
---

`getDocument()` on an editor handle held across `loadDocumentBuffer` returns the document it loaded, with that document's package, instead of the new body under the previous document's package. Loading that result no longer makes the next read throw.
