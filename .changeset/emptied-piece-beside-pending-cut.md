---
"@stll/docx-core": patch
---

Accepting a change that empties one piece of a container a pending insertion split leaves that insertion's cut with the other piece: rejecting the insertion later keeps its range markers in that piece, as rejecting it before the acceptance does.
