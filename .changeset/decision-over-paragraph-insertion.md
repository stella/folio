---
"@stll/docx-core": minor
---

Record the revisions a direct paragraph insertion borders as its sequenced `insertBlocks` effect, so a concurrent revision decision that resolves none of them passes the insertion. Journals that recorded no effect for such an insertion still read and replay; a decision rebased over one is refused.
