---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Give a drawing that carries no relationship one spelling, and classify the preview the second spelling hid. `Image.rId` is a `RelationshipId`, a branded non-empty string the parser mints from a real `r:embed`, `r:id` or `r:link`, so the empty string can no longer stand for absence: it reached a save as `<a:blip r:embed=""/>`. A VML shape's render is now preview-only, like the group render beside it, so the editor declines to manipulate it and a save replays the authored `w:pict` instead of writing the render into `word/media/` as the picture the shape had become. Stored collaboration snapshots take both changes through attr-schema version 4.
