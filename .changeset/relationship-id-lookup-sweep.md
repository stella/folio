---
"@stll/folio-core": patch
---

Every relationship-id lookup goes through the typed resolver. A hyperlink whose `r:id` names a non-hyperlink relationship no longer resolves to that part's path; a `w:headerReference` or `w:footerReference` with no `r:id` is dropped at parse instead of entering the model as an empty id and being written back as the schema-invalid `r:id=""`; a header or footer part with no `.rels` of its own resolves against nothing rather than against the document's relationships, so a part-local `rId1` can no longer name a body part; and a VML `v:imagedata` or grouped picture with no id records absence rather than an empty string. A body reference whose id names nothing is reported in the parse warnings, so a dangling reference is distinguishable from one the author never wrote.
