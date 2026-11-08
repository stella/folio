---
"@stll/folio-core": patch
---

A comment on `insertAfterBlock`, `insertBeforeBlock` or `deleteBlock` now applies in the experimental canonical session where the default engine puts it. On an insertion it covers the inserted paragraphs. On a tracked deletion it covers the deleted words. An insertion with no text and a direct deletion skip, as in the default engine.
