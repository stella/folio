---
"@stll/docx-core": patch
---

Creating a numbering instance puts the instance, and the abstract numbering it brings, before the first entry of a greater id instead of at the end, so concurrent creations leave one order whichever applies first. A creation that restores a captured numbering part, such as undoing a deletion, puts them back where they were: undoing the deletion of an instance that was not the last one is no longer refused.
