---
"@stll/folio-core": minor
---

Add the `@stll/folio-core/text-shaping` entry: the text shaper, the sfnt reader and glyph-id subsetter, and script segmentation. `getShaper({ wasm })` instantiates the shaper from WebAssembly bytes instead of fetching it, and `shaper.resolveBidi` resolves the Unicode Bidirectional Algorithm (levels and visual order, with isolates) for one line.
