import {
  configureFolioMessageAssets,
  type FolioMessageAssetOptions,
  FolioMessageCatalogError,
  getLoadedFolioMessages,
  loadFolioMessages,
  preloadFolioMessages,
} from "./loadMessages";
import { messages as english } from "./messages/en.gen";

export { isFolioLocale } from "./locales";
export { FOLIO_LOCALES, type FolioLocale } from "./messages/locales";
export {
  configureFolioMessageAssets,
  FolioMessageCatalogError,
  loadFolioMessages,
  preloadFolioMessages,
  type FolioMessageAssetOptions,
};

// The exported type is self-contained: declarations reference no JSON asset.
type FolioMessageTree = { [key: string]: string | FolioMessageTree };
export type FolioMessages = { folio: FolioMessageTree };

/**
 * Read the requested catalog after it has loaded. English is always available;
 * an unloaded or failed translated catalog falls back to English. Await
 * `loadFolioMessages` or `preloadFolioMessages` before mounting translated UI.
 */
export const getFolioMessages = (locale: string): FolioMessages =>
  getLoadedFolioMessages(locale) ?? english;
