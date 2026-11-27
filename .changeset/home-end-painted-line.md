---
"@stll/folio-core": patch
---

Home and End move the caret to the start or end of the line as it is painted on the page, in the editor's state as the key is handled. A key pressed right after a click no longer loses its move when the editor's selection is written back before the browser reports the caret it moved. Shift extends the selection, and a vertical step after Home keeps the line-start column.
