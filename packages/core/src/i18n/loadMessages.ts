import { dataFileUnder, readDataFile } from "../utils/packageDataFile";
import { resolveFolioLocale } from "./locales";
import type { FolioMessages } from "./messages";
import { messages as english } from "./messages/en.gen";
import type { FolioLocale } from "./messages/locales";

/**
 * English ships with the editor. Every other catalog is a data file beside the
 * package (`assets/i18n/<locale>.json`, exported as
 * `@stll/folio-core/i18n/catalogs/<locale>.json`): its strings in the order
 * the English catalog holds its keys, fetched the first time the locale is
 * asked for, so no bundle of the editor carries a translated catalog.
 */
type TranslatedLocale = Exclude<FolioLocale, "en">;

export type LocaleCatalogOptions = {
  /**
   * The absolute URL of a directory holding the package's
   * `i18n/catalogs/<locale>.json` files, to fetch from instead of beside the
   * package (a CDN path of the same version, say).
   */
  readonly baseUrl?: string | URL;
  /**
   * Catalogs the host already holds, used without fetching anything: the
   * `messages` of `@stll/folio-core/i18n/messages/<locale>`, for offline
   * hosts or pages whose Content Security Policy blocks the fetch.
   */
  readonly catalogs?: Readonly<
    Partial<Record<TranslatedLocale, FolioMessages>>
  >;
};

let catalogOptions: LocaleCatalogOptions = {};

// One literal `new URL` per locale, so bundlers that copy referenced files
// emit each as an asset. The path holds from `src/` and `dist/` alike.
const packagedCatalogUrl = (locale: TranslatedLocale): URL => {
  switch (locale) {
    case "ar":
      return new URL("../../assets/i18n/ar.json", import.meta.url);
    case "cs":
      return new URL("../../assets/i18n/cs.json", import.meta.url);
    case "de":
      return new URL("../../assets/i18n/de.json", import.meta.url);
    case "es":
      return new URL("../../assets/i18n/es.json", import.meta.url);
    case "et":
      return new URL("../../assets/i18n/et.json", import.meta.url);
    case "fr":
      return new URL("../../assets/i18n/fr.json", import.meta.url);
    case "he":
      return new URL("../../assets/i18n/he.json", import.meta.url);
    case "hi":
      return new URL("../../assets/i18n/hi.json", import.meta.url);
    case "hu":
      return new URL("../../assets/i18n/hu.json", import.meta.url);
    case "lt":
      return new URL("../../assets/i18n/lt.json", import.meta.url);
    case "lv":
      return new URL("../../assets/i18n/lv.json", import.meta.url);
    case "pl":
      return new URL("../../assets/i18n/pl.json", import.meta.url);
    case "pt-BR":
      return new URL("../../assets/i18n/pt-BR.json", import.meta.url);
    case "sk":
      return new URL("../../assets/i18n/sk.json", import.meta.url);
    case "tr":
      return new URL("../../assets/i18n/tr.json", import.meta.url);
    case "zh-CN":
      return new URL("../../assets/i18n/zh-CN.json", import.meta.url);
    default:
      locale satisfies never;
      return locale;
  }
};

const catalogUrl = (locale: TranslatedLocale): URL => {
  const { baseUrl } = catalogOptions;
  return baseUrl === undefined
    ? packagedCatalogUrl(locale)
    : dataFileUnder(baseUrl, `${locale}.json`);
};

const isStringList = (value: unknown): value is readonly string[] =>
  Array.isArray(value) && value.every((item) => typeof item === "string");

const fetchCatalog = async (
  locale: TranslatedLocale,
): Promise<FolioMessages> => {
  const url = catalogUrl(locale);
  const strings = await readDataFile(url);
  if (!isStringList(strings)) {
    throw new Error(`${url.href} is not a folio locale catalog`);
  }
  // Rebuilt from the English shape, which loads with the first catalog.
  const { localeMessages } = await import("./localeMessages");
  return localeMessages(strings);
};

const catalogs = new Map<
  TranslatedLocale,
  FolioMessages | Promise<FolioMessages>
>();

/**
 * Set where translated catalogs come from; each call replaces the previous
 * options. Call it before mounting an editor in a translated locale.
 */
export const configureLocaleCatalogs = (
  options: LocaleCatalogOptions,
): void => {
  catalogOptions = options;
  // A catalog that failed to load is tried again under the new options.
  for (const [locale, messages] of catalogs) {
    if (messages === english) catalogs.delete(locale);
  }
};

/** Read English or a catalog that has already loaded or was handed over. */
export const getLoadedFolioMessages = (
  locale: string,
): FolioMessages | undefined => {
  const resolved = resolveFolioLocale(locale);
  if (resolved === "en") return english;
  const supplied = catalogOptions.catalogs?.[resolved];
  if (supplied !== undefined) return supplied;
  const cached = catalogs.get(resolved);
  return cached instanceof Promise ? undefined : cached;
};

/**
 * Load just the catalog for the requested locale. The load that fails rejects,
 * so its caller can report it; from then on the locale reads as English
 * (editors fall back to it) until {@link configureLocaleCatalogs} is called,
 * rather than every editor fetching and reporting it again.
 */
export const loadFolioMessages = (locale: string): Promise<FolioMessages> => {
  const resolved = resolveFolioLocale(locale);
  if (resolved === "en") return Promise.resolve(english);
  const supplied = catalogOptions.catalogs?.[resolved];
  if (supplied !== undefined) return Promise.resolve(supplied);
  const cached = catalogs.get(resolved);
  if (cached !== undefined) return Promise.resolve(cached);

  const pending = fetchCatalog(resolved).then(
    (messages) => {
      catalogs.set(resolved, messages);
      return messages;
    },
    (error: unknown) => {
      catalogs.set(resolved, english);
      throw error;
    },
  );
  catalogs.set(resolved, pending);
  return pending;
};

/** Test seam: forget every loaded catalog and the options. */
export const resetLocaleCatalogs = (): void => {
  catalogOptions = {};
  catalogs.clear();
};

/** The English catalog, which every editor falls back to. */
export const englishFolioMessages: FolioMessages = english;
