---
"@stll/folio-core": minor
---

Draw a `PreviewDescriptor` instead of rasterising one. A diagram's shapes become filled `rect` primitives in the image's box, so the DOM backend paints the absolutely positioned divs it paints every rectangle as and the PDF exporter emits vector operators rather than embedding a bitmap of a vector drawing. `ImageTable` no longer interns a descriptor, and `MAX_BUILD_PREVIEW_PIXELS` is gone with the rasters it bounded. A shape now keeps the colour the drawing authored, including a channel of zero, which the raster read as the backdrop's.
