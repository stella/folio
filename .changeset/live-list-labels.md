---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

List labels follow the document as it stands. After an operation or edit inserts, deletes, re-levels or restarts list items, `getContent()`, the AI snapshot, `getContentAsText()`, the `read_document` rows and `toMarkdown()` of the edited document show the numbers the page shows, instead of the numbers read when the document was opened. Readers count with the same counter that paints the page's markers.
