---
"@stll/folio-core": patch
---

Removing a hyperlink also removes its Hyperlink style, and a tracked removal records that style as the runs' previous formatting; an edit operation that deletes a note reference takes the note with it, as the editors do: tracked, the note's text is deleted with the reference, and removed outright, the note is dropped on save.
