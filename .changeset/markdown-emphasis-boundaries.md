---
"@stll/folio-core": patch
---

Markdown export writes emphasis that reads back at any boundary: adjacent emphasized runs share their delimiters, delimiters beside punctuation or spaces are placed where they can open and close, and literal tildes and edge underscores are escaped, so no stray `*`, `_` or `~` reaches the text.
