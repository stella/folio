---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Edits report which paragraph a removed one joined: `TouchedBlocks.absorbed` in docx-core results, `paragraphIdentity` on tracked operations whose paragraph joined away, and `absorbedBlocks` on a canonical reviewer's `acceptAll()` / `rejectAll()`. A join keeps the second paragraph's id, except where the joining paragraph's mark records `survivor: "first"`, which keeps the first.
