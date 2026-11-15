---
"@stll/folio-core": patch
---

The layout instrumentation hook reports each selective save attempt through `onSelectiveSaveAttempt({ written, structural })`, so browser measurements can tell a selective text patch, a structural splice and the full repack apart. The hook is undefined in normal use.
