import { FOLIO_LOCALES, type FolioLocale } from "./messages/locales";

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

export const resolveFolioLocale = (locale: string): FolioLocale => {
  if (isFolioLocale(locale)) return locale;
  const base = baseLanguageOf(locale);
  return base !== null && isFolioLocale(base) ? base : "en";
};
