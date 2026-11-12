---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A canonical session's own document is typed read-only (`ReadonlyDocument`) wherever it is handed out: the session's `document`, the hidden editor manager's `readCanonicalDocument`, and the `getDocument`, `getDocumentContext` and `readCanonicalDocument` callbacks the hidden editor, note and header and footer managers read, including the `readCanonicalDocument` props of React's `NoteStoryEditor`, `HiddenHeaderFooterPMs` and `HiddenProseMirror` ref. Readers inside the editor read it through `readModel`. `createCanonicalSession` takes a model, which it copies before it seeds the session. The Vue editor passes its model to these callbacks as it holds it.
