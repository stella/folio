---
"@stll/folio-core": patch
---

A tracked or suggested `deleteTableRow` (and `deleteTable`) over a row holding a pending insertion now marks the inserted text deleted too (`w:ins > w:del`), as `deleteBlock` does inside a paragraph. `getContent()` and every other reader used to keep listing the inserted text of a row that accepting removes.
