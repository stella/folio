---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Accepting a deleted paragraph mark, or rejecting an inserted one, now leaves the paragraph whose mark stays, with its identity and properties; tracked merges carry the first paragraph's properties onto it so accepting still reads as the direct merge. Suggesting mode records joins, cuts, splits, pastes and typing over select-all the same way, so accepting or rejecting them leaves the paragraphs a direct edit or no edit would. A paragraph whose mark goes right before a table runs on into the table's first cell. Clean Markdown (`docxToMarkdown`, `toMarkdown`) accepts paragraph marks the same way. A note follows its reference: rejecting the reference's deletion restores the note's text, undoing that restores the deletion, and accepting it removes the note.
