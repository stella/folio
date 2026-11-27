---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

TIFF images are decoded in a worker loaded through `new URL(…, import.meta.url)` the first time a document holds one, so the decoder is no longer part of the editor's bundle. When decoding cannot run (the worker cannot load, or the image is over the pixel cap) the editor paints a placeholder box in the image's frame, and the document keeps the TIFF.

Hosts that cannot load the worker (offline use, a Content Security Policy without `worker-src`, a single-file CDN bundle) call `configureTiffDecoding({ decoder: await import("utif2") })` from `@stll/folio-core/tiff-decoding`, also exported by `@stll/folio-react` and `@stll/folio-vue`, to decode on the calling thread.
