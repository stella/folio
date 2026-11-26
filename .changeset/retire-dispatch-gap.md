---
"@stll/folio-core": minor
---

Refusals that carried the `unclassified-transactions` gap now name their condition, such as a stale projection, a request the document does not allow or an internal failure, and that gap is retired. A plugin or raw dispatch that changes the document outside document operations is refused under the new `hostTransaction` condition, naming the transaction's meta keys. `groundForDocumentOpRefusal(reason)` returns the ground a document operation refusal reports, and a suggested change the tracked-change records cannot express is refused under the new `tracked-change-records` gap.
