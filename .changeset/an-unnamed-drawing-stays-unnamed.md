---
"@stll/folio-core": patch
---

Read `name=""` on a drawing object as no name. `@name` is schema-required on
`CT_NonVisualDrawingProps`, so a shape, text box or picture the model never
named still writes one, and the reader took that empty string back as authored
content: `absent → save → parse` landed on `""` instead of absent, and the next
save carried it. Nothing downstream can tell the two apart, so the reader now
maps the one value the writer mints back to absence. `@descr` and `@title` are
optional and written only when authored, so `""` in either stays a string
someone wrote.
