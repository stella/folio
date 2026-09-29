---
"@stll/folio-core": patch
---

Accepting a deleted paragraph mark, or rejecting an inserted one, now leaves the paragraph whose mark stays, with its identity and properties; tracked merges carry the first paragraph's properties onto it so accepting still reads as the direct merge. Accepting or rejecting every change also removes a note whose reference it removed.
