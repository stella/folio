---
"@stll/folio-core": patch
---

Keep a paragraph's `w:hyperlink` elements in source order on a full save. A run that follows a hyperlink and carries its own identity (a revision-session attribute, a `w:rPr` it must keep, or mixed content such as a tab and text) no longer saves ahead of the hyperlink, so a paragraph that mixes links and plain runs no longer reads in a different order after it is opened and saved.
