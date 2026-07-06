---
"@stll/folio-core": minor
"@stll/folio-vue": patch
---

Add the UI catalog keys the Vue adapter references but the flat `folio` catalog
was missing, so components render readable English labels instead of raw
key-paths (and stop logging `IntlError: MISSING_MESSAGE`).

Adds 320 keys across the `alignment`, `colorPicker`, `common`, `contextMenu`,
`dialogs` (hyperlink / insertTable / footnoteProperties / tableProperties /
pageSetup / findReplace / insertSymbol / imageProperties / imagePosition),
`editor`, `font`, `formattingBar`, `imageOverlay`, `imageTransform`, `imageWrap`,
`lineSpacing`, `revisions`, `styles`, `table` (+ `table.styles`), `tableAdvanced`,
`toolbar`, `trackedChanges`, `viewer`, and `zoom` families, plus flat
`decreaseFontSize` / `increaseFontSize`. English source values are placeholders in
the 16 non-English locales (grandfathered in the i18n-check baseline); real
translations are a separate concern.

The existing flat `fontSize` label key blocks a `fontSize.*` namespace, so the
Vue toolbar's font-size step buttons resolve `decreaseFontSize` /
`increaseFontSize` instead.
