---
"@stll/folio-core": patch
---

A `"suggested"` `deleteBlock` no longer reaches the saved package: its proposed paragraph-mark deletion was written as a real tracked `w:del` on the paragraph mark, so the saved document carried a join nobody had accepted. Suggested paragraph-mark revisions are now stripped at serialization like every other suggestion.
