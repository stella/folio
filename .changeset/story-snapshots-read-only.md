---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The note and header/footer editor managers' `snapshotDocument` takes and returns the model typed read-only: the snapshot shares what no edit touched with the model. React's `PagedEditorRef.getDocument` is typed read-only for the same reason, and the document React's `getDocument` hands the host owns its notes and numbering.
