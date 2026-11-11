---
"@stll/folio-vue": patch
---

Typing no longer builds a copy of the whole document on every key to refresh the toolbar, the rulers, the page setup controls and the watermark: they read the editor's document model, so a keystroke in a long document costs far less. Embedded fonts load once per loaded document instead of following each settled edit.
