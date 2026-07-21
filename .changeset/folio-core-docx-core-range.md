---
"@stll/folio-core": patch
---

Correct the `@stll/docx-core` dependency range: folio-core imports `normalizeRevisionId`, which only exists in docx-core 0.5.0, but 0.14.0 was published declaring `^0.4.0` (excludes 0.5.0 under 0.x semver). Republish so the range resolves to `^0.5.0`.
