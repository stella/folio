---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Concurrent tracked changes of one paragraph that meet are accepted rather than refused: a cut at a tracked change's edge or inside its range, a direct deletion reaching one of its edges, and overlapping tracked deletions, where the deletion sequenced first holds the shared content. The authority's order decides the document and every peer replays to it; `trackedOrderPair` names these pair kinds. A peer that applied its own edit of such a pair first gets the document as a `replaced` change, which now carries `moves`, the rebased change, when the two documents hold the same paragraphs at the same lengths. A canonical session moves its selection over `moves` instead of falling back to the nearest position.
