---
"@stll/folio-core": patch
---

Saving a document whose main part had no relationships part now creates one when a comment, a reply or another relationship is added. The comment used to be written but not linked, so other applications did not show it, and adding a relationship failed.
