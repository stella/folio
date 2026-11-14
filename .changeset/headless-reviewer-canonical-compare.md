---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` resolves the final section's tracked property changes (`acceptChange`, `rejectChange`, `acceptAll`, `rejectAll`) as a section properties operation in the session commit, and resolves pending suggestions staged in headers, footers and notes in one commit across every story they reach.

Its comparison access reads the session, and stages the target's final section properties, styles, numbering and missing header or footer parts as session commits; staging a terminal table's review carrier, mapped section boundaries and inline atom matching still refuse with the `publicOps.headlessSession` gap.
