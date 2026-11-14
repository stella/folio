---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The context menu's Cut, Delete, Delete image and Paste as plain text work in a canonical session in both adapters: they run `replaceRangeText` (a new core command) or the plain-text paste through the view's command owner, so the session journals them as text input and records them as suggestions while suggesting.
