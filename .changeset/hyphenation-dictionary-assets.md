---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Hyphenation dictionaries now ship as data files (`@stll/folio-core/hyphenation/<id>.json`) that are fetched the first time a document hyphenates in their language, instead of as modules, so no bundle of the editor carries them. Bundlers that copy files referenced through `new URL(…, import.meta.url)` (Vite, webpack) emit them; on Node they are read from the installed package. The editor lays the document out again when a dictionary arrives, and does not hyphenate when one cannot be loaded, reporting the failure to `onError`.

Hosts that cannot fetch them from beside the package (a single-file CDN bundle, offline use, a Content Security Policy that blocks the request) call `configureHyphenationAssets` from `@stll/folio-core/hyphenation`, also exported by `@stll/folio-react` and `@stll/folio-vue`. It takes a `baseUrl` holding the same files, or the parsed files as `dictionaries`. `preloadHyphenationDictionaries` loads dictionaries ahead of a layout that must hyphenate on its first pass.
