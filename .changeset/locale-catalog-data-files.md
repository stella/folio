---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Translated UI catalogs now ship as UTF-8 data files (`@stll/folio-core/i18n/catalogs/<locale>.json`) that the editor fetches the first time a locale is used, instead of as modules, so no bundle of the editor carries them; English still ships with the editor. Mounting and a locale switch wait for the catalog before painting translated UI, as before. When a catalog cannot be loaded, the editor opens in English and reports the failure to `onError`, instead of failing to mount.

Hosts that cannot fetch the files from beside the package (offline use, a Content Security Policy that blocks the request, a single-file CDN bundle) call `configureLocaleCatalogs` from `@stll/folio-core/i18n/messages` (also `@stll/folio-react/messages` and `@stll/folio-vue/messages`) with a `baseUrl` holding the same files, or with `catalogs` taken from the `@stll/folio-core/i18n/messages/<locale>` modules. `loadFolioMessages` loads a catalog ahead of time.
