---
"@stll/folio-core": patch
---

Harden the paraId lifecycle: duplicate resolution in the allocator now maps the original paragraph's position through the transaction (pasting a copy above its source no longer steals the source's id), allocation transactions are excluded from paragraph change tracking, and the hex id generators can no longer mint the reserved `00000000` value.
