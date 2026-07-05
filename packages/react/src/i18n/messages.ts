// folio's UI translation catalog now lives framework-neutrally in
// `@stll/folio-core/i18n/messages` so non-React adapters (Vue, Nuxt) can consume
// the same single source of truth without pulling in React. This module keeps
// the `@stll/folio-react/messages` public subpath stable by re-exporting it.
export * from "@stll/folio-core/i18n/messages";
