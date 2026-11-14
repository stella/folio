---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Rejecting a change keeps every bookmark and move range boundary its content held unless the boundary is a copy that change owns, whatever else the story holds. Accepting a copy keeps it, as inserting it directly does. So rejecting a boundary's insertion and accepting a copy of it leave the same boundaries in either order, in the engine and in the editor. Comment boundaries, which record no copies, still go where the story holds them otherwise.
