---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Rejecting an insertion that holds another reviewer's pending deletion of a bookmark or other range boundary now leaves that deletion pending, holding only the boundary, where the insertion stood. Accepting the deletion later removes the boundary, as deleting it directly does, and rejecting it keeps the boundary, so the result is the same whichever change is resolved first, in the engine and in the editor. A nested deletion holding no boundary still goes with the insertion.
