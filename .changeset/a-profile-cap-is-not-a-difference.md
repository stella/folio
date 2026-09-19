---
"@stll/folio-core": patch
---

Keep content identity across the content-structure profile caps. A table past 128 blocks, or past the retained-text budget, had its exact signature, anchor texts and token counts blanked, and the fallback pairing read the absence as evidence: the only table in a document came back deleted and re-inserted when the document was compared with itself. The profile now carries a digest of every block at any size, so identical content pairs before any heuristic runs, and the heuristics it could not compute are modelled as `skipped-over-cap` rather than as empty.
