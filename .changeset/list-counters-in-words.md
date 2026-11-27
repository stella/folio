---
"@stll/folio-core": patch
---

Template-plan labels write a list's `cardinalText` and `ordinalText` counters in words when the paragraph's language is English (`One`, `Twenty-first`, `One hundred one`), as the reference consumer prints them, so such a label can be written as text. In a language whose words are not known the label still cannot be written as text.
