---
"@stll/docx-core": patch
---

Rejoin link and content control pieces around kept range boundaries in any rejection order: count kept boundaries a pending deletion holds one by one, let a pending deletion holding only kept boundaries hold the seam between the pieces, and read past rejected new content joined to nothing when placing the boundaries.
