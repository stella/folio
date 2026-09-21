---
"@stll/folio-core": patch
---

Keep a tracked move's `w:name` through the editor.

`w:moveFromRangeStart` and `w:moveToRangeStart` carry the name that binds a move's source to its destination, and with it the range's own `w:id`, author and date. `toProseDoc` had no node for the four move-range markers and dropped them, so the first save after any edit wrote two unrelated revisions where the document had one relocation.

The markers are not content a caret can sit in, so they ride on the paragraph beside `bookmarks` and are put back around the wrappers they delimit: a range opens before the first `w:moveFrom` / `w:moveTo` of its kind and closes after the last, and a range that spans several paragraphs still opens in the first and closes in the last. What is carried is the model's own marker, so a field added to `CT_MoveBookmark` is carried without being named again.
