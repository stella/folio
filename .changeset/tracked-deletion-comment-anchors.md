---
"@stll/docx-core": patch
---

A tracked deletion over a record that holds a comment anchor, such as a content control or hyperlink holding a comment's end, is refused as a deletion of a bare anchor is. Accepting such a deletion kept the anchor that deleting directly removes, so the two gave different documents. A planned deletion still deletes around the anchor.
