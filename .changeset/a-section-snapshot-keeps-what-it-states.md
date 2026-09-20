---
"@stll/folio-core": minor
---

Keep a `w:lnNumType` or `w:pgBorders` that states nothing, on the live section and inside a `w:sectPrChange`.

Both types declare every attribute and every child optional, so the element alone is a legal document: it says the section is line numbered, or bordered, with the defaults. The parser recorded it and the two serializers wrote nothing for it, which `serializeDocGrid` in the same file had already decided the other way — an attribute-less element is a document folio must write back, and the record exists only because the parser read one.

The loss showed up differently in the two places the element can sit. On a live section the record then held a field the serializer did not write, so the package fidelity guard refused the save outright. Inside a `w:sectPrChange` the refusal is swallowed and the change is written with an empty `<w:sectPr/>`, so the snapshot a reviewer would restore came back blank.
