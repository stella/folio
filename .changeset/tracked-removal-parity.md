---
"@stll/folio-core": patch
---

Removing a hyperlink also removes its Hyperlink style, and a tracked removal records that style as the runs' previous formatting; a tracked note reference deletion deletes the note's content with it, and a note whose reference is gone is dropped on save.
