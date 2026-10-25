---
"@stll/docx-core": patch
---

When rejecting leaves a kept range start after its kept end, the boundary kept from the later revision moves, so the earlier one stays where it was whatever order the revisions are rejected in.
