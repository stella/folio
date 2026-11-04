---
"@stll/folio-core": patch
---

Saving from the editor writes `w:softHyphen` and `w:noBreakHyphen` back as elements instead of turning them into characters inside `w:t`; the same characters authored as text stay text.
