---
"@stll/folio-vue": patch
---

Fix the published subpath exports so consumers resolve real artifacts: the
`./composables`, `./dialogs`, and `./styles` `types` conditions pointed at
`dist/*.d.ts` files the build never emits (the declarations flatten to
`dist/composables/index.d.ts`, `dist/components/dialogs/index.d.ts`, and
`dist/styles/index.d.ts`), and `./messages` pointed at a `./src/*.ts` source
file absent from the published tarball. `./messages` is now a built entry
(`dist/messages.js` / `.cjs` + `dist/i18n/messages.d.ts`).
