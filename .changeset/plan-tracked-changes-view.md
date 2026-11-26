---
"@stll/folio-core": minor
---

A template plan can be made for one view of a document with tracked changes: its text with every change accepted, or with every change rejected (`trackedChanges: "accept" | "reject"` on the plan, and as an option of `analyzeDocxBlocks`). The blocks, labels and hashes are that view's, the output carries no revision, verification holds it to that view (`revision-left` otherwise), and the report says what was resolved. A paragraph mark the view removes joins its paragraph to the next; a row the view removes goes whole; rejected, recorded formatting comes back. A manifest read without a view says what the document tracks (`revisions`: counts by kind, authors, first and last date). Without a view on the plan such a source is refused as before.
