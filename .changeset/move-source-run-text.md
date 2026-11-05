---
"@stll/folio-core": patch
---

A tracked move source (`w:moveFrom`) is saved with its run text as `w:t` and its field code as `w:instrText`, as ISO/IEC 29500-1 §17.3.3.7 and §17.13.5.22 specify; `w:delText` and `w:delInstrText` stay reserved for runs inside a `w:del`. Move sources read with `w:delText` load as before and save with `w:t`.
