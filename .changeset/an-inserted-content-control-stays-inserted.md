---
"@stll/folio-core": minor
---

Keep the revision that covers a whole inline content control.

`w:ins > w:sdt` reached the editor as a control whose leaves carried nothing. The control is an `inline*` node rather than an inline atom, so it is not a run carrier and the revision mark had nowhere to land; the save leg then wrote the control back beside the revision that had held it, and text a reviewer had inserted was no longer inserted.

The revision now rides the leaves the control holds, and a revision that covers all of them is written back around the control. Accepting or rejecting it is an operation over the control itself: rejecting an inserted control removes it rather than leaving an empty one standing where it was, on the editor-command path and on the headless one alike. A revision over part of the content has no such form and stays where the editor holds it, per child.

`w:ins > w:sdt` and `w:sdt > w:ins` reach the editor as the same marks on the same leaves, so a span rebuilt from the editor comes back revision-outermost, as it already does for a hyperlink and for a transparent wrapper. A paragraph nobody edited keeps its authored order, because selective save replays its bytes.
