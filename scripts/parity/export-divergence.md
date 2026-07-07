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
