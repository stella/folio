---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A drawing with no picture relationship keeps the markup it arrived with. A `w:drawing` whose graphic is a chart or an OLE frame, or which carries no `a:graphic` at all, has no `a:blip` and so no relationship id; `Image.rId` is now absent in that case rather than an empty string, and a save writes the anchor back as authored instead of rebuilding it into a picture bound to whichever relationship the part happens to list first. Relationship ids resolve through one typed resolver that distinguishes a resolved id from an absent and a dangling one, and an image reference that names a non-image relationship no longer resolves to that part.
