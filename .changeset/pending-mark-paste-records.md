---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

A paragraph inserted directly after one whose paragraph mark is a pending insertion now goes in where rejecting that insertion leaves the same paragraphs as inserting into the rejected document: the pending break moves to the last new paragraph. A suggested paste of a paragraph whose break is pending deletion now records the paste's insertion of the break with the copied deletion beside it, which stays pending once the paste is accepted. A pasted paragraph carrying a table's review carrier now pastes as the plain paragraph it shows, directly or as a suggestion.
