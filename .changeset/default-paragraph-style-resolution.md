---
"@stll/folio-core": patch
---

The default paragraph style is resolved the way ECMA-376 17.7.4.17 does rather than by assuming the style id `Normal`: the paragraph style flagged `w:default="1"` wins, the last one where several are flagged, then a style carrying the built-in `w:name`, then the built-in id. A localized or generated package that names its default `Standard`, `Normln` or `style0` now resolves it, and one that declares no default gets a minted default instead of a failed extraction.
