---
"@stll/folio-core": minor
---

feat(core): segment-fit line-breaking seam for plain-text measurement

Add a swappable `SegmentFitEngine` seam to the layout-engine measurement path.
A registered engine fits plain-text runs from prepared segment widths instead of
the word-walk's per-call word re-measurement and `findMaxFittingLength`
slice-probe binary search. Gated behind the new `segmentFitLineBreaking`
measurement feature flag; with the flag off or no engine installed, the legacy
walk runs byte-identically. `@stll/premirror-bridge` provides a
`@chenglou/pretext`-backed engine with frozen parity tests.
