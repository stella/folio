---
"@stll/docx-core": minor
---

Add tracked changes to `@stll/docx-core/ops`: a `revision` stamp on the text, formatting and paragraph operations records the edit as a tracked insertion, deletion, property change or paragraph mark, with an exact inverse. `planTrackedDeletion` plans a tracked deletion around the author's own insertions, and `revisionIdDemand` counts the revision ids an operation takes. The operation schema is now version 2.
