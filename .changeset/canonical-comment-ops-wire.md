---
"@stll/docx-core": patch
---

Comment operations can be sent in a batch and sequenced: creations (ranges, points, replies and revision comments), content edits, resolutions, deletions and their undo inverses, with bounded validation of the comment records they carry and the author and date they were sent with. A range or point comment's anchor moves like text over concurrent text and structure edits, and concurrent edits move past its reference; concurrent comment operations commute where they use different comments or different fields of one comment, and are refused with `partConflict` naming the new `comments` part otherwise. A new comment takes its place in the comment list by id. `UNSENT_OP_TYPES` is removed: every operation kind is sendable.
