---
"@stll/folio-core": patch
---

Keep an authored image or shape transform through a save, including zero. `rot="0"` and `flipH="0"` are OOXML's defaults, so the truthiness guard on `a:xfrm` could not tell an authored zero from an absent attribute and the save dropped it; the image parser also read `rot="0"` as no rotation at all. Rotation and both flips are now read as authored values and written iff the model holds one.
