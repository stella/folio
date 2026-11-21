---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Report the pending property changes an operation folds its edit into. A tracked docx-core operation lists them in the new `AppliedDocumentOp.folded`, beside the new ids in `revisions`, and `combineRecordedRevisions` combines both lists over operations applied in turn. A canonical batch receipt names the id of a run or paragraph property change its operation now shares, so accepting or rejecting the receipt's ids covers the operation's edit.
