---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Remove the verbatim sink's attribute remainder. `PreservedMarkup.attributes`, the `PreservedAttribute` type, the dispatcher's `modelsAttribute` option and `serializePreservedAttributes` had no caller in the product: no container ever passed the predicate, so no attribute was ever kept, and the shape read as coverage that was not there. `docs/container-contract.md` records the design and what wiring it needs.
