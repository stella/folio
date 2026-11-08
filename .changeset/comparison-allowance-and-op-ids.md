---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

`generateRedlineDocx` aligns every story of one comparison against a single alignment allowance, as `compareDocxVersions` does. The reference sequencer replays a stored outcome only to the actor that first used its operation id, and the optimistic client refuses a broadcast that reuses a pending operation id under another actor.
