import ar from "./messages/ar.json";
import cs from "./messages/cs.json";
import de from "./messages/de.json";
import en from "./messages/en.json";
import es from "./messages/es.json";
import et from "./messages/et.json";
import fr from "./messages/fr.json";
import hu from "./messages/hu.json";
import lt from "./messages/lt.json";
import lv from "./messages/lv.json";
import pl from "./messages/pl.json";
import ptBR from "./messages/pt-BR.json";
import sk from "./messages/sk.json";

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
  "hu",
  "lt",
  "lv",
  "pl",
  "pt-BR",
  "sk",
] as const;

export type FolioLocale = (typeof FOLIO_LOCALES)[number];

// The exported type is self-contained: it references neither `typeof en` nor
// any locale JSON, so the emitted `messages.d.ts` carries no `./messages/*.json`
// import. dist inlines the JSON into the JS chunk and ships no JSON files, so
// such an import would be unresolvable for a published TypeScript consumer of
// `@stll/folio-react/messages` (Codex #11). The contract is just a mergeable
// single-namespace message object.
type FolioMessageTree = { [key: string]: string | FolioMessageTree };
export type FolioMessages = { folio: FolioMessageTree };

// `Widen` stays internal. Against the English catalog (the structural source of
// truth) it keeps the exact key shape, so every locale is checked to carry the
// same keys, while relaxing the per-key literal values to `string`; that is what
// makes the 13 translated catalogs assignable to one shared type.
type Widen<T> = { [K in keyof T]: T[K] extends string ? string : Widen<T[K]> };

const CATALOGS: Record<FolioLocale, Widen<typeof en>> = {
  en,
  de,
  fr,
  es,
  cs,
  ar,
  et,
  hu,
  lt,
  lv,
  pl,
  "pt-BR": ptBR,
  sk,
};

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
