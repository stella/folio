---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Refuse a tracked insertion after a container's final paragraph whose mark is under review with the new `pendingParagraphMarkDeletion` or `pendingParagraphMarkInsertion` refusal reason, by what the mark has pending, and a tracked join only into a table under review as `joinIntoPendingTable` (any other table it cannot run into stays `untrackable`). The canonical session maps compiler refusals through one total table: `pendingBreakTakeover` reports `pendingParagraphMarkDeletion`, each pending mark reason reports the public reason of its name, and only reasons for shapes the engine does not record yet report `unsupportedBlock` with a capability gap.
