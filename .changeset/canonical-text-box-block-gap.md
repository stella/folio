---
"@stll/folio-core": patch
---

A public operation in a canonical session that names a paragraph the session holds no address in, such as one in a text box, now refuses as an unsupported block under the unsupported public payload gap instead of reporting the block as missing. Every refused operation in a batch now names its gap, including one resolved again after the block deletion it was paired with was refused.
