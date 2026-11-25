import { resolveFolioLocale } from "./locales";
// The generated TS catalogs keep raw JSON out of the source-mirrored build.
// Import a messages/<locale> entry to load one catalog independently.
import { CATALOGS as GENERATED_CATALOGS } from "./messages/catalogs.gen";
import type { FolioLocale } from "./messages/locales";

export { loadFolioMessages } from "./loadMessages";
export { isFolioLocale } from "./locales";
export { FOLIO_LOCALES, type FolioLocale } from "./messages/locales";

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

const CATALOGS: Record<
  FolioLocale,
  Widen<(typeof GENERATED_CATALOGS)["en"]>
> = GENERATED_CATALOGS;

// Returns folio's bundled `{ folio: ... }` messages for `locale`. An exact
// match wins; otherwise a regional tag resolves to its base-language catalog
// (`de-DE` -> `de`) — a host usually passes the same tag it gives IntlProvider —
// before falling back to English for anything folio does not ship (`pt-BR` is a
// shipped locale, so the exact match above already covered it). Merge into app
// messages so the app wins on the rare intentional override:
//   { ...getFolioMessages(locale), ...appMessages[locale] }
export const getFolioMessages = (locale: string): FolioMessages =>
  CATALOGS[resolveFolioLocale(locale)];
