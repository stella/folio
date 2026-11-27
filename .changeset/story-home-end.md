---
"@stll/folio-core": patch
---

Home and End in header, footer and note editors move the caret to the start or end of the rendered line on every platform, and Shift extends the selection. On some platforms the browser's own Home and End only scrolled and left the caret where it was. The move is set in the editor's state as the key is handled.
