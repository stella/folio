---
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-vue": patch
---

Read the item before an index without wrapping: at a list's start there is no previous item, where `.at(index - 1)` read the list's last one.
