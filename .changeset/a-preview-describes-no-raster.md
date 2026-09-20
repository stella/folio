---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Retire the preview rasteriser. `PreviewDescriptor` drops `pixelWidth` and `pixelHeight`, which sized a raster nothing builds any more, and `previewRaster.ts` goes with them; `MAX_PREVIEW_SHAPES` moves to the diagram reader, which is what applies it.
