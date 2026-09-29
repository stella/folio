---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

Accepting a deleted paragraph mark, or rejecting an inserted one, now leaves the paragraph whose mark stays, with its identity and properties; tracked merges carry the first paragraph's properties onto it so accepting still reads as the direct merge. Suggesting mode records joins, cuts, splits, pastes and typing over select-all the same way, so accepting or rejecting them leaves the paragraphs a direct edit or no edit would. A note follows its reference: rejecting the reference's deletion restores the note's text, and accepting it removes the note.
