---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A replacement over whole paragraphs that hold pending insertions, words or breaks, retires the comments they cover, as it does over paragraphs without them: a direct edit removes any author's insertion with its paragraph, and a suggestion retracts its own author's. In a canonical session, typing over everything after a suggested paste beside a comment applies instead of being refused.
