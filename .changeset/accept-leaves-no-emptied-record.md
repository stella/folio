---
"@stll/docx-core": patch
---

Accepting tracked changes one operation at a time now gives the same result as accepting them at once: a record an acceptance empties goes, as a direct deletion leaves none, instead of staying empty beside the piece of it a later acceptance brings back. Rejecting a paste that cut a content control, hyperlink or other inline container now gives the container back whole: the piece the rejection empties folds into the other piece even when no resolved change lies between them.
