---
"@stll/folio-core": patch
---

The bundled font families, their directories, the measured subsets and weights are declared once, in `fonts/bundledFonts`, which the resolver, the editor's availability check and the `@stll/folio-fonts` generator all read. `FONTSOURCE_PACKAGES` is still exported from `fonts/fontsourceFaces`.
