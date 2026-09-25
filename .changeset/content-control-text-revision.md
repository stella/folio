---
"@stll/folio-core": patch
---

Accepting a tracked deletion of all of an inline content control's text, or rejecting an insertion of it, now leaves the control in place, emptied, and the saved redline keeps such a revision inside `w:sdtContent`. Only a revision around the control (`w:ins > w:sdt`, `w:del > w:sdt`), a tracked deletion spanning the whole control in the editor, or a deleted paragraph removes the control with it.
