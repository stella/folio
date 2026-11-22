---
"@stll/docx-core": patch
---

Resolving tracked changes beside rejected insertions leaves the same document in any order. What was made inside a rejected insertion that joined nothing keeps no join past it, and takes only the range markers kept from its own record. A cut goes with the insertion it was made in and is read between its own pieces. Range markers kept at a seam go where the seam goes: into the record beyond the insertions it waited on, or into the one an accepted deletion's pieces were cut from. A rejected split's cut goes on past an empty paragraph split off in it later, and an insertion at the edge of a split's piece is part of that piece.
