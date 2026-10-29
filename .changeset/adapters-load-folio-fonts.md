---
"@stll/folio-react": patch
"@stll/folio-vue": patch
"@stll/folio-cli": patch
---

The bundled document fonts now come from `@stll/folio-fonts` instead of eight `@fontsource` packages, so an application installs folio's packages only. The stylesheets declare the same families, weights, styles and unicode ranges as before, as `.woff2` alone, and the CLI measures and renders with the same `.woff` files.
