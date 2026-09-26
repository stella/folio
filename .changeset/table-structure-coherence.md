---
"@stll/folio-core": patch
---

Deleting a table column directly and accepting the same deletion tracked now leave the same table. Where the deleted column held a row's only cell of its own beside a vertical merge, the direct deletion left a row of nothing but that merge: the saved package wrote it as a row of `w:vMerge` continuations, which reopens with the merge split apart, so the reopened snapshot, `getContent()` and `docxToMarkdown` disagreed on its spans. Such a row is now removed and the merge closes over one row fewer, as accepting the tracked deletion already did. Accepting a tracked deletion of a cell that spans several rows now removes the cell instead of moving it into the row below.
