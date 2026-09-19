---
"@stll/folio-core": patch
---

Read a `w:val` outside `ST_OnOff` as absent in both shapes of the type. The element shape (`<w:b w:val="yes"/>`) used to read anything it could not parse as `true` while the attribute shape read the same malformed value as nothing, so one type answered the same input two ways. A census of 5,316 public documents from 289 producers found no producer writing a non-standard spelling systematically, so no tolerance beyond `1`, `0`, `true`, `false`, `on` and `off` is added.
