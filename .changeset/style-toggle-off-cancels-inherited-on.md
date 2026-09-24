---
"@stll/folio-core": patch
---

Resolve a toggle run property's `basedOn` chain by nearest-defined-wins: a style's own explicit off now cancels a toggle inherited from its base style, instead of being ignored.
