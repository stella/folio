---
"@stll/folio-core": patch
---

A point comment (a comment reference with no range) is no longer lost when the text before it is deleted and the deletion accepted: it stays at its reference, as rejecting keeps it, and saves back as a point instead of a range over that text.
