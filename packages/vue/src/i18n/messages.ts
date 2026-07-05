// folio's UI translation catalog lives framework-neutrally in
// `@stll/folio-core/i18n/messages` (the single source of truth shared with the
// React adapter). This keeps the `@stll/folio-vue/messages` public subpath
// stable by re-exporting it.
export * from "@stll/folio-core/i18n/messages";
