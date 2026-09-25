---
"@stll/folio-core": patch
---

Honour `w:rFonts/@w:hint="eastAsia"`: the Latin-1 symbols, General Punctuation through Dingbats, Greek, Cyrillic, spacing modifiers and private-use characters of a hinted run measure and paint with its East Asian font. Basic Latin and accented Latin letters keep the `w:ascii`/`w:hAnsi` font, and `w:cs`/`w:rtl` runs are unaffected.
