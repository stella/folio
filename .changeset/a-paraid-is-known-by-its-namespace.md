---
"@stll/folio-core": patch
---

Read a `paraId` by namespace URI wherever it is written. A paragraph's id, a `commentsExtensible` join key and a `commentsExtended` thread link were each resolved by prefix, with a local-name fallback that matched any prefix at all: a file binding `w14`, `w15` or `w16cex` to a prefix of its own was read correctly only by luck, and an unrelated `vendor:paraId` was read as a thread key, carrying another thread's date, parent and resolved state into the comment. One reader now answers "the paraId of this element" for all of them, and it accepts the Word 2010, 2012 and 2018 namespaces and no others.
