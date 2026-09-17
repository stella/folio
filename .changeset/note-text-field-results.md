---
"@stll/folio-core": patch
---

Read a field in note, header and footer text as its stored result rather than the text the editor paints for it. A field with no result used to contribute a synthesized placeholder, and a DATE field the current date, so the same document produced different text on different days and anything hashing or comparing that text was unstable. Every other inline atom contributes exactly what it did before.
