---
"@stll/docx-core": patch
---

Resolving tracked changes at a cut record leaves the same document in more orders. Insertions made one at another's edge at the end of a cut piece are part of that piece, as deep as they are joined to it, and the one nearest the seam hands its join on when it goes. A cut goes on through the pieces its own change removes, and the range markers kept there go into the records it leaves. Markers kept from a record of a rejected insertion go into the content made inside that record, and an accepted deletion that takes a split's piece takes the split's cut with it.
