---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A tracked merge's paragraph property change belongs to the merge: accepting or rejecting the merge's paragraph-mark deletion decides the property change too, on both engines and after a save and reopen. The change keeps its own revision id and names the deletion it belongs to (`ParagraphPropertyChange.resolvesWith`, saved as `folio:resolvesWith`; review history version 2, version 1 still read). A tracked merge's `revisionId` is its paragraph-mark deletion.
