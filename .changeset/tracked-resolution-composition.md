---
"@stll/folio-core": patch
---

Resolving tracked edits that build on pending ones is exact. `rejectAll` no longer throws after a tracked merge into an inserted paragraph, and gives the original document back: removing a paragraph break that is itself a pending insertion (a tracked merge, or deleting a paragraph whose break a split or an insertion added) retracts that insertion instead of overwriting it, and deleting a paragraph that is wholly a pending insertion removes it. Rejecting a tracked split with a tracked table inserted between its halves joins them again. A comment keeps covering one stretch of text through tracked replacements, merges and inserted paragraphs or tables, so its anchored text reads the same before and after a save. Accepting every suggestion keeps every suggested paragraph in the saved document even when one of them could not first become a tracked change.
