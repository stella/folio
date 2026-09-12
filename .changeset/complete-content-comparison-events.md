---
"@stll/folio-core": minor
---

Add an ordered, representation-neutral `compareContent` API that captures declared caller input into a canonical owned comparison result with typed text, formatting, structural, and move events. Canonical paragraph insertion and source-removal boundaries preserve container identity through transport lowering, including terminal moves whose predecessor carries the removed break. Split and merge events include canonical text segments and paragraph formatting. Document-version comparison now consumes the same comparison core through one bounded session across all package stories and reports aggregate resource-limit failures through `FolioVersionComparisonLimitError`; comparison executors remain package-private. AI block preview runs now expose complete `effectiveFormatting` and `authoredFormatting` values in the document model's native units; the flattened formatting fields and `directFormatting` alias are removed.
