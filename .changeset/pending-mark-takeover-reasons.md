---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Refuse a tracked insertion after a container's final paragraph whose mark is under review with the new `pendingParagraphMark` refusal reason, and a tracked join only into a table under review as `joinIntoPendingTable` (any other table it cannot run into stays `untrackable`). The canonical session maps compiler refusals through one total table: `pendingBreakTakeover` and `pendingParagraphMark` report `pendingParagraphMarkDeletion`, and only reasons for shapes the engine does not record yet report `unsupportedBlock` with a capability gap.
