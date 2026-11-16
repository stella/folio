---
"@stll/folio-core": patch
---

Refuse canonical replacements whose aligned comment anchors cannot be preserved, before changing document state or history. Pure deletions in surviving content follow ranged-comment retirement; replacing a whole span of mixed comment coverage refuses rather than returning inconsistent comment definitions across edit modes.
