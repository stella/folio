---
"@stll/folio-vue": patch
---

`scrollToParaId` finds a paragraph on a page that is not painted yet: it looks the paragraph up in the document, scrolls to it, and flashes it once its page is painted, as the React editor does.
