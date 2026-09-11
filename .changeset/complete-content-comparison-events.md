---
"@stll/folio-core": minor
---

Add an ordered, representation-neutral `compareContent` API with canonical owned inputs and typed text, formatting, structural, and move events. Split and merge events include canonical text segments and paragraph formatting. Document-version comparison now consumes the same comparison core through one bounded session across all package stories and reports aggregate resource-limit failures through `FolioVersionComparisonLimitError`; comparison executors remain package-private.
