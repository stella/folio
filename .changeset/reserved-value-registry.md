---
"@stll/docx-core": patch
---

Record, next to each model type, what its fields decide about the reserved values of the OOXML slots they carry (`w:numId` 0, `w:outlineLvl` 9, a `w:tcW` number under `w:type="auto"`, and so on). Each disposition map is total over its type by construction, so a model field added without a recorded decision fails `bun run typecheck`.
