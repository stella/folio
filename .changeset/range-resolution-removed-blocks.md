---
"@stll/folio-core": patch
---

Accepting or rejecting the tracked changes of a range resolves a paragraph mark, row or cell only while it still stands: one inside a table the same resolution removes no longer acts on the block after that table, so a range covering the whole document resolves it as accepting or rejecting all changes does.
