# Intentional React <-> Vue export divergence

`scripts/check-export-parity.mjs` fails when `@stll/folio-react` and
`@stll/folio-vue` drift on either their `package.json` `exports` subpaths or
their `src/index.ts` named exports. Every divergence that is intentional (a
framework-native difference, or a not-yet-ported surface) must be listed here as
a first-backtick list item. The gate reads only the first backticked token on
each `-` list line, so prose backticks do not widen the opt-out.

## Subpath divergences (`package.json` `exports`)

React ships a standalone CSS bundle; the Vue adapter exposes its advanced
surfaces (UI primitives, composables, dialogs) as dedicated subpaths that have
no React equivalent.

- `./compat/eigenpal` — React-only migration entrypoint for archived Eigenpal users.
- `./standalone.css` — React-only standalone stylesheet.
- `./ui` — Vue-only UI-primitive subpath.
- `./composables` — Vue-only composables subpath (React ships hooks from root).
- `./dialogs` — Vue-only dialog-component subpath.
- `./styles` — Vue-only style entry.

## Named-export divergences (`src/index.ts`)

Known Vue gaps (React-only chrome/components not yet ported to Vue):

- `AutocompleteCaretOverlay` — React overlay component; no Vue equivalent yet.
- `AutocompleteCaretOverlayProps` — type for the React-only overlay.
- `AutocompleteCaretRect` — type for the React-only overlay.
- `clampZoom` — standalone zoom util re-exported only from the React hook module.
- `formatZoom` — standalone zoom util (React-only re-export).
- `parseZoom` — standalone zoom util (React-only re-export).
- `ZOOM_PRESETS` — standalone zoom preset table (React-only re-export).
- `UseWheelZoomOptions` — React `useWheelZoom` hook option type.
- `UseWheelZoomReturn` — React `useWheelZoom` hook return type.

Vue-native runtime (no React equivalent — React consumes `use-intl` directly):

- `defaultLocale` — Vue i18n runtime.
- `i18nPlugin` — Vue i18n plugin install fn.
- `provideLocale` — Vue i18n provide helper.
- `useTranslation` — Vue i18n composable.

Per-dialog prop/data type aliases (framework-idiom difference, not a port gap):
React's dialogs are `.tsx` components that export a named props/data type alias
per dialog from root, so React hosts can type their dialog usage. Vue's dialogs
are `.vue` SFCs whose props are consumed through the component itself and carry
no equivalent root-exported type alias. The dialog _components_ are paired
(both adapters export `FindReplaceDialog`, `WatermarkDialog`, etc. from root);
only these React type aliases have no Vue counterpart.

- `FindReplaceDialogProps` — React find/replace dialog props; Vue SFC props are component-internal.
- `FootnotePropertiesDialogProps` — React footnote-properties dialog props.
- `HyperlinkBookmarkOption` — React hyperlink dialog bookmark-option type.
- `HyperlinkDialogData` — React hyperlink dialog data type.
- `HyperlinkDialogProps` — React hyperlink dialog props.
- `ImagePositionData` — React image-position dialog data type.
- `ImagePositionDialogProps` — React image-position dialog props.
- `ImagePropertiesData` — React image-properties dialog data type.
- `ImagePropertiesDialogProps` — React image-properties dialog props.
- `InsertImageDialogData` — React insert-image dialog data type.
- `InsertImageDialogProps` — React insert-image dialog props.
- `InsertTableDialogData` — React insert-table dialog data type.
- `InsertTableDialogProps` — React insert-table dialog props.
- `InsertTableStyleOption` — React insert-table style-option type.
- `PageSetupDialogProps` — React page-setup dialog props.
- `PasteSpecialDialogProps` — React paste-special dialog props.
- `PasteSpecialMode` — React paste-special mode union (Vue keeps the equivalent union inside its SFC).
- `SplitCellDialogData` — React split-cell dialog data type.
- `SplitCellDialogProps` — React split-cell dialog props.
- `TableProperties` — React table-properties dialog model type (Vue keeps the equivalent in `./ui`).
- `TablePropertiesDialogProps` — React table-properties dialog props.
- `WatermarkDialogProps` — React watermark dialog props.

Known React gaps (Vue-only chrome not yet ported to React):

- `InsertSymbolDialog` — Vue insert-symbol dialog component; React has no insert-symbol dialog yet.
