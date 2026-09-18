---
"@stll/folio-core": patch
---

A template preview value carrying newlines now breaks its lines inside the marker's paragraph instead of painting them on top of the content below it: each newline becomes a line break run, the way `w:br` does, so the paragraph measures the height its value needs.
