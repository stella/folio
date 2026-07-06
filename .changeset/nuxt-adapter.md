---
"@stll/folio-nuxt": minor
---

Add `@stll/folio-nuxt`, a Nuxt 3 & 4 module wrapping `@stll/folio-vue`. It
auto-imports an SSR-safe, client-only `<DocxEditor>` component and the public Vue
composables, injects the editor stylesheet, and forces the editor packages
through Vite's dependency optimizer. Configurable via the `docxEditor` key
(`prefix`, `injectStyles`).
