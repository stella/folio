---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Canonical session refusals now name the gap they fall under: a pasted table refuses under the tracked table structure gap; a pasted embedded block, a pasted comment or note reference, and a composition over content the session cannot project refuse under the story content gap; typing in a text box refuses under the unsupported public payload gap, as a public operation on its paragraphs does. A host plugin's transaction that changes the document is refused with a message naming the transaction's meta keys, and the `experimentalSession` prop documents that such edits are expressed as document operations.
