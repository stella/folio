---
"@stll/folio-core": patch
---

Keep the words a link or a field nested inside one of its own kind puts on the page.

`CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`, so either may hold another of itself. folio does not model the nesting — Word writes neither, and the public corpus has four nested links in one package written by a converter — and captured the inner element through the child sink, which knows nothing about a capture beyond its bytes. So the markup survived and the text did not: a link inside a link, and a field inside a field's cached result, reached the editor as an opaque atom showing nothing, and text extraction, markdown and layout skipped the words entirely.

Both are captured through the element instead, the way `w:customXml` and `w:smartTag` already were, so the capture carries the visible text beside the markup. Position is unchanged: the capture is still a member of the owner's own content union, between the same two children it was read between.
