---
"@stll/folio-react": patch
"@stll/folio-core": patch
---

Every user-facing string in the React editor now goes through use-intl. ~173 hardcoded English JSX literals (concentrated in the dialogs: image position/properties, page setup, footnotes, watermark, tables, hyperlinks, header/footer editing) were wired to the locale catalogs — many onto existing keys the components weren't using, ~50 new keys added to en.json and synced to all locales. Visible English output is unchanged. A new `no-untranslated-jsx-literal` oxlint rule enforces this in CI so untranslatable copy cannot land again.
