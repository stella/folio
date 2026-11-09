---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Typing in the experimental canonical session no longer copies the whole document after each edit to bring the header, footer and note editors up to date: they read the session's document, which each edit replaces. A change notification now makes one copy for the host instead of three.
