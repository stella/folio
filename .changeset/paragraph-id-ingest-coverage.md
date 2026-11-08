---
"@stll/folio-core": patch
---

`ensureParaIds` replaces a `w14:paraId` that is not eight hex digits (the parser reads it as no id) and stamps header, footer, and note parts wherever the document relationships place them. `readBilingualDocx` refuses a row whose handle is a package id several body paragraphs share.
