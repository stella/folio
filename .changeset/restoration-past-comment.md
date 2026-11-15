---
"@stll/docx-core": patch
---

A paragraph restoration, such as undoing an edit, and a concurrent new comment anchored in that paragraph now meet by the restoration's span: when neither end of the comment's anchor lies inside the stretch the restoration changes, both are kept in either order instead of one being refused because the other changed the paragraph.
