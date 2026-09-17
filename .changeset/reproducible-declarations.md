---
"@stll/folio-core": patch
---

Name the element type of two derived constants so the emitted declarations stop varying between builds. Their inferred type was a union the declaration emitter wrote out member by member in an order that changed from build to build, which made `dist/**/*.d.ts` non-reproducible.
