---
"@stll/folio-core": patch
---

A paragraph whose own break is a pending insertion no longer shows a struck list marker for numbering a suggestion removed from it: it had no number before the change. The page painted a number while editing and the bare level text (`%1.`) after reopening the file; it now paints none in both.
