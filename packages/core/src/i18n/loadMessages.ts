import { resolveFolioLocale } from "./locales";
import type { FolioMessages } from "./messages";
import { messages as english } from "./messages/en.gen";
import type { FolioLocale } from "./messages/locales";

const loaders = {
  ar: () => import("./messages/ar.gen"),
  cs: () => import("./messages/cs.gen"),
  de: () => import("./messages/de.gen"),
  es: () => import("./messages/es.gen"),
  et: () => import("./messages/et.gen"),
  fr: () => import("./messages/fr.gen"),
  he: () => import("./messages/he.gen"),
  hi: () => import("./messages/hi.gen"),
  hu: () => import("./messages/hu.gen"),
  lt: () => import("./messages/lt.gen"),
  lv: () => import("./messages/lv.gen"),
  pl: () => import("./messages/pl.gen"),
  "pt-BR": () => import("./messages/pt-BR.gen"),
  sk: () => import("./messages/sk.gen"),
  tr: () => import("./messages/tr.gen"),
  "zh-CN": () => import("./messages/zh-CN.gen"),
} satisfies Record<
  Exclude<FolioLocale, "en">,
  () => Promise<{ messages: FolioMessages }>
>;

const catalogs = new Map<
  Exclude<FolioLocale, "en">,
  FolioMessages | Promise<FolioMessages>
>();

/** Read English or a catalog that has already loaded. */
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

  const pending = loaders[resolved]().then(
    ({ messages }) => {
      catalogs.set(resolved, messages);
      return messages;
    },
    (error: unknown) => {
      catalogs.delete(resolved);
      throw error;
    },
  );
  catalogs.set(resolved, pending);
  return pending;
};
