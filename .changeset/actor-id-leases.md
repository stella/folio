---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

The reference sequencer leases each actor disjoint blocks of paragraph, revision, content-control and comment ids and rejects a batch that introduces an id outside its actor's lease. Id allocators and the canonical session mint only from a lease when they hold one; without one they keep minting the lowest free ids.
