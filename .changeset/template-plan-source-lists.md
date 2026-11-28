---
"@stll/folio-core": minor
---

A template plan may keep the source's own list for a paragraph: `numbering: { numId: "new", ilvl, sourceNumId }`. The output carries a copy of the source's list instance and of its definition (a definition that only links a numbering style is followed to the one that defines its levels), numbered after the template's own; every paragraph naming the same `sourceNumId` counts in that copy, as the source counts them, and the level overrides the source's instance states go with it. What ties a definition to its package (its ids, the paragraph styles its levels name, picture bullets) stays behind, and the copy declares the namespaces the source's numbering part binds. Verification holds every kept list to a definition the output holds; the labels it prints are compared with the source's like any other.
