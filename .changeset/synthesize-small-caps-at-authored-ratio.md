---
"@stll/folio-core": patch
---

Measure and paint `w:smallCaps` text at a fixed 0.8-of-run-size synthesized-capital ratio instead of a browser's uncontrollable `font-variant: small-caps`, so a run's lowercase letters no longer draw (and measure) too small and shift line wraps. `w:caps` still wins over `w:smallCaps` when both are set.
