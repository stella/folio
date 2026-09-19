---
"@stll/folio-core": patch
---

Keep every comment's author, body, date, resolved state and reply link with its own `w:id` across a save. `word/comments.xml` was written with the top-level comments first and the replies after, so a document whose comments.xml interleaves a reply with a later thread root came back from the next parse in a different `comments[]` order than it went in, and anything reading that array by position saw one comment's text and author under another's place. Comments are now written in the model's order, both comment parts are planned once from one ordered list of ids, and a duplicate or missing `w14:paraId` resolves to the same comment on parse and on save.
