---
"@stll/docx-core": patch
---

Text replacing words inside a hyperlink, at its start or end, or typed between two of its leaves, stays in the one link, edited directly or tracked. A tracked insertion can continue a record whose piece the pending deletion before it holds; accepting the tracked replacement gives the direct edit in either resolution order, and rejecting it gives the source back. Bookmark and comment markers at the link's edges keep the link on the new text.
