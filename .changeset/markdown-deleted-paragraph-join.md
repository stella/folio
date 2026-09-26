---
"@stll/folio-core": patch
---

`docxToMarkdown` with `trackedChanges: "clean"` no longer hands a wholly deleted paragraph's style or list to the paragraph below it: a tracked deletion of the last bullet of a list rendered the following body paragraph as a bullet, while accepting the same change in the editor or through `FolioDocxReviewer.acceptAll()` leaves it a plain paragraph.
