---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` resolves the final section's tracked property changes (`acceptChange`, `rejectChange`, `acceptAll`, `rejectAll`) as a section properties operation in the session commit, and resolves pending suggestions staged in headers, footers and notes in one commit across every story they reach.

Its comparison access reads the session, and stages the target's final section properties, styles, numbering, missing header or footer parts, mapped section boundaries and a paired table's own properties as session commits. Staging a review carrier after a terminal table refuses with the `story-content-projection` gap, and inline atom and provenance matching still refuse with the `publicOps.headlessSession` gap.
