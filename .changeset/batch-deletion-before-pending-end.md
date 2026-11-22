---
"@stll/folio-core": patch
---

A batch on the canonical session that inserts a paragraph and deletes the paragraphs after it, where only pending deletions follow, retires the inserted paragraph's break as sending the operations one at a time does, so accepting it keeps the paragraph before the insertion apart.
