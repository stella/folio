---
"@stll/folio-core": patch
---

i18n quality gate and terminology glossary. A new `i18n-lint` stage in `check:i18n` verifies every translated catalog string for placeholder parity with the English source, ICU MessageFormat validity, CLDR plural-category completeness per locale, dropped plurals/exact selectors, and glossary terminology (with a ratchet baseline for future debt; the current catalogs are clean). New `glossary.json` term base mimics Microsoft Word's localized terminology (LibreOffice divergences documented), covering 46 word-processing concepts across 16 locales with forbidden nonstandard variants; three existing translations were corrected to canonical Word terms (pt-BR "Recortar", tr "Açıklama").
