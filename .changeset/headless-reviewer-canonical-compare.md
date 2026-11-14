---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` resolves the final section's tracked property changes (`acceptChange`, `rejectChange`, `acceptAll`, `rejectAll`) as a section properties operation in the session commit, and resolves pending suggestions staged in headers, footers and notes in one commit across every story they reach.
