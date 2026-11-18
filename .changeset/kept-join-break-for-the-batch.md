---
"@stll/folio-core": patch
---

Keep a pending join's break as soon as a tracked deletion of the paragraph it runs into applies, so later operations in the same batch read the joined paragraph as they do after that deletion alone.
