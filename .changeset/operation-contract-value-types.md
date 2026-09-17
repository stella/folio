---
"@stll/folio-core": patch
---

Name the element type of the operation contract's alignment and line-spacing value lists. Spreading a const tuple widened them to an array of the union, and the declaration emitter wrote that union out member by member in an order that changed between builds, so the emitted declarations were not reproducible.
