---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Open a `w:sym` that names only one of its two optional attributes. `CT_Sym` declares `w:font` and `w:char` optional and Word renders a `<w:sym w:char="F0B7"/>` by falling back to the run's font; folio refused the document twice over, once in the model validator and once in the ProseMirror projection, so a file Word opens did not open at all. Both checks now accept an absent attribute and still reject a malformed character, and the serializer writes an absent attribute back as absent instead of inventing `w:font=""`.
