---
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The React and Vue editors no longer install the ProseMirror suggestion-mode plugin or read the paragraph change tracker: the canonical session owns the editing mode and reports pending changes.
