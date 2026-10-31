---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A batch now equals its operations applied one at a time with every operation that addresses a block before any operation that deletes or merges that block away. `@stll/docx-core/ops` exports the shared order (`orderAddressingBeforeRemoval`) and the rule for which deletions take a block out (`deletionRemovesBlock`). In tracked changes, properties a batch sets on a paragraph it merges away, or on the last paragraph it deletes, no longer survive the carried properties: accepting gives what the same batch applied directly gives.
