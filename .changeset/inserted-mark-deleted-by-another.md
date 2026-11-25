---
"@stll/folio-core": patch
---

A paragraph whose inserted mark another author deleted reads as pending deletion, so a `replaceBlock` there is refused (`pendingDeletion`) instead of running its text into the next paragraph once accepted.
