---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

A document-operation batch no longer applies an operation to a target an earlier operation of the same batch already claims. Every operation addresses the document as it was read, so an edit inside a paragraph the batch also deletes, rewrites, splits or merges, or an edit of text another operation already changes, landed on positions the first one had moved: a replacement followed by a deletion of the same paragraph rewrote the next paragraph instead. The later operation is now refused before anything is applied, with the new skip reason and issue code `overlappingOperation` (`retryable: true`, `recovery: "refreshDocument"`) and a `message` naming the earlier operation; `suggest_changes` explains it in words. Comments and formatting around an edit still apply with it, and so do disjoint edits of one paragraph, a chain of merges and, while tracked, a merge into a paragraph the batch deletes. Setting the properties of a paragraph a batch inserts a paragraph before is no longer skipped as `missingBlock`, and a row insertion is placed through the positions the batch has already changed.
