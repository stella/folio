---
"@stll/folio-core": patch
---

Match a container's declared children by namespace as well as local name in the shared child dispatcher. A child from another namespace now reaches the sink instead of the handler its local name happens to collide with, so `m:r` is no longer read as a text run, emptied and pruned.
