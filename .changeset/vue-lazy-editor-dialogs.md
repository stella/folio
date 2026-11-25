---
"@stll/folio-vue": patch
---

The Vue editor loads its dialogs (find and replace, hyperlink, symbol, image and table properties, page setup, numbering value, watermark) the first time each opens, and mounts each only while it is open, as the React editor does. A dialog mounted already open now sets itself up as one opened later does.
