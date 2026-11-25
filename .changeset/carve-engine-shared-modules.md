---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
"@stll/folio-agents": patch
---

Shared pieces of the ProseMirror editing engine move to their own folio-core modules: paragraph-id allocation (`prosemirror/paraIds`), the change tracker's transaction meta (`prosemirror/paragraphChangeTrackerMeta`, `prosemirror/documentStructureCounts`), the history commands and `HistoryShortcutOwner` (`prosemirror/extensions/core/historyCommands`), the AI edit applier's public types and table builders (`ai-edits/applyTypes`, `ai-edits/tableNodes`), and the document-operation applier (`document-operations-apply`). Package entry points export the same names.
