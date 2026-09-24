---
"@stll/folio-core": patch
---

Keep the reserved `xml:` prefix on attributes such as `xml:space` and `xml:lang` when a save rebuilds an element the model has no field for, instead of writing them back unprefixed.
