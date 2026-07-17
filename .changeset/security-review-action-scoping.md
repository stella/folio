---
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Scope review actions to their intended target. AI `undoDocumentOperations`
always undoes the body operation instead of whatever story (header/footer)
currently has focus; sidebar accept/reject applies only the selected revision
rather than every tracked change in the resolved range (React and Vue); the Vue
page-setup dialog respects read-only; Vue table properties target the active
editor view; and Vue AI-authored comments use the requested operation author.
