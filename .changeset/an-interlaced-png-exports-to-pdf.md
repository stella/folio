---
"@stll/folio-core": patch
---

Export a document containing an interlaced PNG to PDF. The PDF image decoder refused Adam7 outright, alongside the 16-bit samples it cannot represent exactly, so a document Word displays without comment failed to export at all. Adam7 is lossless and exactly representable — seven ordinary filtered rasters of the same samples — so it is now decoded: each pass is unfiltered and scattered into the full raster, and the decompression budget counts the passes rather than assuming progressive geometry. A 16-bit PNG is still refused, interlaced or not.
