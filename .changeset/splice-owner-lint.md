---
"@stll/folio-core": patch
---

Route the remaining part patchers through the splice owner: stamping paragraph
ids and restoring a numbering level's custom format both cut regions out of a
serialized part by hand, so a comment range crossing one of those regions could
lose a half. A lint rule now holds the boundary.
