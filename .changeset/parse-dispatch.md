---
"@stll/folio-core": patch
---

Parse DOCX files faster: child dispatch reads handler tables directly, run and paragraph property handlers are built once, and every XML part uses the streaming parser.
