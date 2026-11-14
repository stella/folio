---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Table menu, table toolbar and context menu actions, Insert Table, page break insertion, and the Vue numbering value and table properties dialogs run their editor commands through the view's command owner, so they work on the canonical engine: the session takes each command's intent instead of refusing the view's transaction. A lint rule holds every package source that still runs a command on a view's own state and dispatch to a shrinking count.
