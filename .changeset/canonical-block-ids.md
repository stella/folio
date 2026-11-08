---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Block-level public operations in the experimental canonical session now leave paragraph ids where the default engine leaves them, so host references stay valid across engines. Inserting after the last paragraph of a list and deleting that paragraph directly now apply instead of refusing. In the engine, `INSERT_BLOCKS` now records a tracked insertion after a paragraph that no paragraph follows, and a tracked insertion before a table. A new `insertParagraphs` editor intent compiles to `INSERT_BLOCKS`, and `joinParagraphs` takes a `survivor` for direct joins. A paragraph mark gains an in-memory `survivor` that is never written to a file. Rejecting a suggested insertion at the end of a list gives the document back exactly.
