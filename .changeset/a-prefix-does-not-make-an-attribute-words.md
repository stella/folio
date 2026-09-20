---
"@stll/folio-core": patch
---

Read a paragraph's `w14:textId` by namespace URI. A prefix is an alias, and the prefix lookup fell through to an any-prefix local-name match, so a `textId` bound to a foreign namespace was taken for Word's paragraph identity and written back as one. An alternate prefix bound to the Word 2010 URI still reads.
