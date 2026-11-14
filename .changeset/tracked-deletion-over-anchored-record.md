---
"@stll/docx-core": patch
---

A tracked deletion of a range takes in a record it covers whole that holds a comment anchor, such as a content control or hyperlink holding a comment's end. Accepting the deletion removes the anchors it holds, as deleting the range directly does, so the two give the same document; rejecting it gives the document back. A bare comment anchor is still left standing by a planned deletion, which deletes around it.
