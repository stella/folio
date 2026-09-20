---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Describe a SmartArt diagram at parse time instead of rasterising it. `Image.preview` carries a bounded `PreviewDescriptor`, and the display list builds the PNG when it interns one, so a package's parse no longer pays megabytes per diagram for a picture nothing may paint. The rendered picture is unchanged.
