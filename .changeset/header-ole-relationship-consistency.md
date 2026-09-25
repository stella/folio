---
"@stll/folio-core": patch
---

Keep every relationship id a part's XML still references in sync with what a save prunes from that part's own `.rels`. When package reconciliation drops a relationship whose target the output does not hold (an embedded object a header still names, for example), it now also removes the matching `r:id`-style attribute from the part's content instead of leaving an id nothing resolves. A new `checkPackageIntegrity` helper verifies both invariants — every `r:id` resolves in its own part's `.rels`, and every internal relationship names a part the package holds — for tests to assert against.
