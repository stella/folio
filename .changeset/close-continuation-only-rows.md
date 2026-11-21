---
"@stll/docx-core": patch
---

A merge or column deletion that leaves a row nothing it starts, only vertical merges from above, now removes the row, directly and on accepting a tracked column deletion. A tracked merge records that row's deletion beside its merge records, so accepting gives the direct result and rejecting the original. A merge whose addressed paragraph it drops puts the caret in the merged cell, and a merge beside a pending row insertion or deletion outside its rows no longer waits for it.
