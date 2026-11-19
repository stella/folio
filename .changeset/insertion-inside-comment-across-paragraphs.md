---
"@stll/folio-core": patch
---

Give text an edit puts at a paragraph's edge, or into a blank paragraph, the comments whose text lies on both sides of it across the paragraph boundary, in direct and tracked mode. A saved comment range covers that text, so the live document now does too: removing the comment's first stretch afterwards leaves the same range whether or not the document was saved in between.
