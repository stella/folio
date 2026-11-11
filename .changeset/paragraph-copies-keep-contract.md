---
"@stll/docx-core": patch
---

`withParagraphCopies` puts copies of a story's paragraphs, at any depth, in their places, its section views included. Each copy must hold the same value under every field as the paragraph it replaces, so a document known to meet the seed contract stays known to meet it and the next operation skips the full check.
