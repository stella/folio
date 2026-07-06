// The per-locale catalogs are inlined into a generated TS module rather than
// imported as raw `./messages/*.json`. Core's tsdown `unbundle: true` build
// mirrors every imported `.json` to a per-locale `dist/i18n/messages/<locale>.js`
// module, for which rolldown emits a malformed
// `export { <locale>_default as default, folio }` (the `folio` named binding is
// never declared). That breaks any downstream bundler processing core's dist.
// Importing from `catalogs.gen.ts` keeps `.json` out of the build graph, so no
// per-locale `.js` is emitted. The `*.json` files remain the source of truth the
// i18n tooling reads; run `bun scripts/i18n-catalogs-gen.ts` to regenerate.
import { CATALOGS as GENERATED_CATALOGS } from "./messages/catalogs.gen";

// folio bundles its own UI translations: the editor reads the `folio.*`
// namespace via `useTranslations("folio")`, and a consumer merges this catalog
// into its app messages so the editor localizes itself with no host-supplied
// strings. The single date-picker label the editor used to read from `common`
// is folded into `folio.*`, so folio owns exactly one top-level namespace and a
// shallow merge against the host's other namespaces never collides.
export const FOLIO_LOCALES = [
  "en",
  "de",
  "fr",
  "es",
  "cs",
  "ar",
  "et",
  "he",
  "hi",
  "hu",
  "lt",
  "lv",
  "pl",
  "pt-BR",
  "sk",
  "tr",
  "zh-CN",
] as const;

export type FolioLocale = (typeof FOLIO_LOCALES)[number];

// The exported type is self-contained: it references neither the generated
// catalog nor any locale JSON, so the emitted `messages.d.ts` carries no
// `./messages/*.json` import. dist inlines the JSON into the JS chunk and ships
// no JSON files, so such an import would be unresolvable for a published
// TypeScript consumer of `@stll/folio-react/messages` (Codex #11). The contract
// is just a mergeable single-namespace message object.
type FolioMessageTree = { [key: string]: string | FolioMessageTree };
export type FolioMessages = { folio: FolioMessageTree };

// `Widen` stays internal. Against the English catalog (the structural source of
// truth) it keeps the exact key shape, so every locale is checked to carry the
// same keys, while relaxing the per-key literal values to `string`; that is what
// makes the 13 translated catalogs assignable to one shared type.
type Widen<T> = { [K in keyof T]: T[K] extends string ? string : Widen<T[K]> };

const CATALOGS: Record<FolioLocale, Widen<(typeof GENERATED_CATALOGS)["en"]>> =
  GENERATED_CATALOGS;

const FOLIO_LOCALE_SET = new Set<string>(FOLIO_LOCALES);

export const isFolioLocale = (locale: string): locale is FolioLocale =>
  FOLIO_LOCALE_SET.has(locale);

// Base language subtag of a BCP-47 tag (`de-DE` -> `de`), or null for a
// structurally invalid tag. `Intl.Locale` throws on malformed input, so this
// boundary parse is guarded.
const baseLanguageOf = (locale: string): string | null => {
  try {
    return new Intl.Locale(locale).language;
  } catch {
    return null;
  }
};

// Returns folio's bundled `{ folio: ... }` messages for `locale`. An exact
// match wins; otherwise a regional tag resolves to its base-language catalog
// (`de-DE` -> `de`) — a host usually passes the same tag it gives IntlProvider —
// before falling back to English for anything folio does not ship (`pt-BR` is a
// shipped locale, so the exact match above already covered it). Merge into app
// messages so the app wins on the rare intentional override:
//   { ...getFolioMessages(locale), ...appMessages[locale] }
export const getFolioMessages = (locale: string): FolioMessages => {
  if (isFolioLocale(locale)) {
    return CATALOGS[locale];
  }
  const base = baseLanguageOf(locale);
  if (base !== null && isFolioLocale(base)) {
    return CATALOGS[base];
  }
  return CATALOGS.en;
};
