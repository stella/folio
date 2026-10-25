---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

When rejecting leaves a kept range start after its kept end, the boundary kept from the later revision moves, in the engine and the editor alike, so the earlier one stays where it was whatever order the revisions are rejected in.
