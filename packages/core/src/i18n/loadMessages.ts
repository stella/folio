import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import { readRuntimeJsonAsset } from "../utils/runtimeJsonAsset";
import { resolveFolioLocale } from "./locales";
import type { FolioMessages } from "./messages";
import { messages as english } from "./messages/en.gen";
import { FOLIO_LOCALES, type FolioLocale } from "./messages/locales";

type TranslatedFolioLocale = Exclude<FolioLocale, "en">;

export type FolioMessageAssetOptions = {
  /** Absolute directory URL holding `i18n/<locale>.json` files. */
  readonly baseUrl?: string | URL;
  /** Catalogs the host already holds, used without fetching. */
  readonly catalogs?: Readonly<
    Partial<Record<TranslatedFolioLocale, FolioMessages>>
  >;
};

const FolioMessageCatalogErrorBase: TaggedErrorClass<"FolioMessageCatalogError"> =
  TaggedError("FolioMessageCatalogError");

export class FolioMessageCatalogError extends FolioMessageCatalogErrorBase<{
  message: string;
  locale: TranslatedFolioLocale;
  cause: unknown;
}> {}

let assetOptions: FolioMessageAssetOptions = {};
const catalogs = new Map<
  TranslatedFolioLocale,
  FolioMessages | Promise<FolioMessages>
>();

// Literal URLs let bundlers copy each data file without putting its strings in
// JavaScript. Construct only the requested URL: an English-only non-ESM host
// must not evaluate `import.meta.url` while its editor module initializes.
const packagedCatalogUrl = (locale: TranslatedFolioLocale): URL => {
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
    default: {
      const exhaustive: never = locale;
      return exhaustive;
    }
  }
};

const catalogUrl = (locale: TranslatedFolioLocale): URL => {
  const { baseUrl } = assetOptions;
  if (baseUrl === undefined) return packagedCatalogUrl(locale);
  const url = new URL(baseUrl);
  if (!url.pathname.endsWith("/")) url.pathname += "/";
  url.pathname += `${locale}.json`;
  return url;
};

const hasCatalogShape = (
  value: unknown,
  shape: unknown,
): value is FolioMessages => {
  if (typeof shape === "string") return typeof value === "string";
  if (
    typeof value !== "object" ||
    value === null ||
    Array.isArray(value) ||
    typeof shape !== "object" ||
    shape === null ||
    Array.isArray(shape)
  ) {
    return false;
  }
  const valueRecord = value as Record<string, unknown>;
  const shapeRecord = shape as Record<string, unknown>;
  const keys = Object.keys(shapeRecord);
  return (
    Object.keys(valueRecord).length === keys.length &&
    keys.every(
      (key) =>
        Object.hasOwn(valueRecord, key) &&
        hasCatalogShape(valueRecord[key], shapeRecord[key]),
    )
  );
};

const loadCatalog = async (
  locale: TranslatedFolioLocale,
): Promise<FolioMessages> => {
  const value =
    assetOptions.catalogs?.[locale] ??
    (await readRuntimeJsonAsset(catalogUrl(locale)));
  if (!hasCatalogShape(value, english))
    throw new Error(
      `The ${locale} catalog does not match the English catalog.`,
    );
  return value;
};

/** Read English or a translated catalog that has already loaded. */
export const getLoadedFolioMessages = (
  locale: string,
): FolioMessages | undefined => {
  const resolved = resolveFolioLocale(locale);
  if (resolved === "en") return english;
  const cached = catalogs.get(resolved);
  return cached instanceof Promise ? undefined : cached;
};

/** Load just the catalog for the requested locale. */
export const loadFolioMessages = (locale: string): Promise<FolioMessages> => {
  const resolved = resolveFolioLocale(locale);
  if (resolved === "en") return Promise.resolve(english);
  const cached = catalogs.get(resolved);
  if (cached !== undefined) return Promise.resolve(cached);

  const pending = loadCatalog(resolved).then(
    (messages) => {
      catalogs.set(resolved, messages);
      return messages;
    },
    (cause: unknown) => {
      catalogs.delete(resolved);
      throw new FolioMessageCatalogError({
        message: `The ${resolved} message catalog could not be loaded.`,
        locale: resolved,
        cause,
      });
    },
  );
  catalogs.set(resolved, pending);
  return pending;
};

/** Load every translated catalog needed by `locales` before mounting UI. */
export const preloadFolioMessages = (
  locales: Iterable<string>,
): Promise<Result<void, FolioMessageCatalogError>> => {
  const requested = new Set<TranslatedFolioLocale>();
  for (const locale of locales) {
    const resolved = resolveFolioLocale(locale);
    if (resolved !== "en") requested.add(resolved);
  }
  return Promise.all([...requested].map(loadFolioMessages)).then(
    () => Result.ok(undefined),
    (error: FolioMessageCatalogError) => Result.err(error),
  );
};

/** Set where translated catalogs come from. Call before mounting editors. */
export const configureFolioMessageAssets = (
  options: FolioMessageAssetOptions,
): void => {
  assetOptions = options;
};

/** Test seam: forget translated catalogs and asset configuration. */
export const resetFolioMessageAssets = (): void => {
  assetOptions = {};
  for (const locale of FOLIO_LOCALES) {
    if (locale !== "en") catalogs.delete(locale);
  }
};
