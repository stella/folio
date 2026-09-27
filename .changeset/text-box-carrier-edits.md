---
"@stll/folio-core": patch
---

A `deleteBlock` on a paragraph that holds a text box now deletes the box with it: directly, and in tracked mode by putting the box's drawing inside the paragraph's deletion, so accept-all removes it and reject-all keeps it. Accept-all and accepting one change now also remove a text box whose drawing was a deleted run. `insertAfterBlock` (and an inserted table or signature table after such a paragraph) now lands after the box's paragraphs, so the reader's block order no longer changes when the package is saved and reopened. A direct `splitBlock` keeps a box with the half that holds it; a tracked split that would leave a box's paragraph apart from the box is refused as `unsupportedBlock`, as a merge of such a paragraph with the next already was.
