---
"@stll/docx-core": patch
---

Planning a tracked edit reads a package's revision, content-control and reserved ids once per package instead of on every attempt, so a tracked replacement in a long document plans about three times faster.
