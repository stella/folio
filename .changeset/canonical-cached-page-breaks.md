---
"@stll/folio-core": patch
---

The experimental canonical session now opens documents that hold cached page-break markers (`w:lastRenderedPageBreak`). Before, any such document was refused and could not be edited in the session. A marker stays where the file put it when text is typed beside it and the document is saved.
