---
"@stll/folio-core": patch
---

A public `replaceBlock` that names a paragraph style restyles the paragraph in the canonical session, direct and tracked, also when its words stay the same. A canonical batch that changes nothing is reported as a no-op without naming a capability gap.
