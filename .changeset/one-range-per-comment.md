---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A comment owns one range. `insertContent` gains `comments` and `newCommentIds` under operation schema 18: an inserted comment anchor of a kind its comment already has standing in the document is a copy, which takes a fresh comment id and a copy of the thread the insertion carries when the slice holds the comment's range whole, and is left out otherwise; anchors completing a range, or of a comment the document does not hold, go in as they are. Direct, tracked and suggested insertions follow it, with exact inverses. Older operation envelopes remain explicitly refused. On the legacy session, pasting commented text the document already holds gives the pasted copy a fresh comment id and a copy of its thread, undone with the paste, or leaves the comment out when the session holds no thread for it.
