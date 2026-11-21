---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Two paragraphs with bookmark or other range boundaries standing between them in the block list can be joined, directly or tracked, and the canonical engine merges them with `mergeBlockWithNext` as the default engine does. Directly, the boundaries go into the joined paragraph where the two meet; tracked, they stay between the blocks and accepting the join gives the direct result. A paragraph that ends a section still needs the next one to touch it. `@stll/docx-core/ops` exports `isBlockBoundary`. A block-list rewrite keeps each range boundary's own record, also beside an equal one.
