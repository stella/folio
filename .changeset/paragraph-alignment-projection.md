---
"@stll/docx-core": minor
---

Project effective paragraph alignment (`w:jc` from direct, style chain, and docDefaults) as a seventh paragraph tuple slot; `DOCX_PROJECTION_SCHEMA_VERSION` moves from 4 to 5. `start`/`end` and other unsupported `w:jc` values project as absent; table-style and numbering-level alignment are not consulted, so a paragraph aligned only by those tiers projects the lower-tier style or docDefaults value.
