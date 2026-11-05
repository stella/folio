---
"@stll/folio-core": patch
---

Rejecting every change gives the document before back where a paragraph runs on into a following table: a table it removes or an inserted paragraph it empties no longer holds the paragraph in place, a deleted first row takes the words, and a deletion nested in an insertion no longer counts as words left. An acceptance also runs past a table it removes.
