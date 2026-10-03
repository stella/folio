// Shipped locale identifiers; this entry does not load translation catalogs.
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
