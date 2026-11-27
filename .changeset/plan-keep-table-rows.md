---
"@stll/folio-core": minor
---

A template plan that keeps its source's tracked changes carries table rows too: an inserted row keeps its record, and a deleted row, which the accepted view does not have, is put back as a deleted row with its text. A deletion made inside an insertion (text one author inserted and another deleted again) is carried with both records. Such sources were refused before; cell changes and moves still are.
