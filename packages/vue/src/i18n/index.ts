/**
 * Vue i18n binding for folio.
 *
 * folio's UI translation catalog is framework-neutral and lives in
 * `@stll/folio-core/i18n/messages` (the same single source of truth the React
 * adapter consumes). This module binds it to Vue with `provide`/`inject` over
 * `use-intl`'s framework-neutral `createTranslator`, so folio's SFCs read the
 * `folio.*` namespace via `useTranslation()` exactly as the React components
 * read it via `useTranslations("folio")`.
 *
 * @packageDocumentation
 * @public
 */
import { createTranslator } from "use-intl/core";
import {
  type App,
  inject,
  type InjectionKey,
  type MaybeRefOrGetter,
  onServerPrefetch,
  provide,
  shallowRef,
  toValue,
  watch,
} from "vue";

import {
  getLoadedFolioMessages,
  loadFolioMessages,
} from "@stll/folio-core/i18n/loadMessages";
import type { FolioMessages } from "@stll/folio-core/i18n/messages";
import { messages as english } from "@stll/folio-core/i18n/messages/en";

/** BCP-47 tag folio falls back to when a host supplies no locale. */
export const defaultLocale = "en";

// Namespaced translator: `t("bold")` resolves `folio.bold`, matching the React
// `useTranslations("folio")` call sites so ported components need no rewrite of
// their `t(...)` calls beyond the catalog keys themselves.
export type FolioTranslator = ReturnType<typeof createFolioTranslator>;

/**
 * Call signature folio's SFCs use: a namespaced `folio.*` key path plus
 * optional ICU values, returning the localized string. The underlying
 * `FolioTranslator` types its key parameter from the catalog literal, which the
 * `createTranslator` overloads collapse to `never` for a widened message type;
 * folio's components address keys as plain paths (matching `presets.ts`), so the
 * public binding exposes this string-keyed signature.
 */
export type TranslateFn = (
  key: string,
  values?: Record<string, unknown>,
) => string;

const createFolioTranslator = (locale: string, messages: FolioMessages) =>
  createTranslator({
    locale,
    messages,
    namespace: "folio",
  });

const LOCALE_KEY: InjectionKey<MaybeRefOrGetter<string>> =
  Symbol("folioLocale");

const fallbackTranslator = shallowRef(
  createFolioTranslator(defaultLocale, english),
);

/**
 * Provide folio's translator to descendant components. `locale` may be a ref,
 * getter, or plain string. An unloaded catalog replaces the current messages
 * once it arrives. Call inside a parent component `setup()`. Before hydrating
 * a non-English editor, preload the catalog with
 * `await loadFolioMessages(locale)` from `@stll/folio-core/i18n/loadMessages`.
 */
export const provideLocale = (
  locale: MaybeRefOrGetter<string> = defaultLocale,
): void => {
  provide(LOCALE_KEY, locale);
};

type LocaleStatus =
  { type: "loading" } | { type: "ready" } | { type: "failed"; error: Error };

const englishTranslation = {
  translator: fallbackTranslator,
  status: shallowRef<LocaleStatus>({ type: "ready" }),
};

const useLocaleTranslator = (
  locale: MaybeRefOrGetter<string>,
  onError?: (error: Error) => void,
) => {
  if (locale === defaultLocale) return englishTranslation;

  const translator = shallowRef(fallbackTranslator.value);
  const status = shallowRef<LocaleStatus>({ type: "loading" });
  let pending: Promise<void> | undefined;
  watch(
    () => toValue(locale),
    (tag, _previous, onCleanup) => {
      const messages = getLoadedFolioMessages(tag);
      if (messages) {
        translator.value = createFolioTranslator(tag, messages);
        status.value = { type: "ready" };
        pending = undefined;
        return;
      }

      const hasTranslation = status.value.type === "ready";
      if (!hasTranslation) status.value = { type: "loading" };
      let active = true;
      onCleanup(() => {
        active = false;
      });
      pending = loadFolioMessages(tag).then(
        (loaded) => {
          if (active) {
            translator.value = createFolioTranslator(tag, loaded);
            status.value = { type: "ready" };
          }
          return undefined;
        },
        (cause: unknown) => {
          if (!active) return undefined;
          const error =
            cause instanceof Error ? cause : new Error(String(cause));
          if (!hasTranslation) status.value = { type: "failed", error };
          if (!onError) throw error;
          onError(error);
          return undefined;
        },
      );
      return pending;
    },
    // Vue disposes non-sync watchers during SSR setup.
    { immediate: true, flush: "sync" },
  );
  onServerPrefetch(() => pending);
  return { translator, status };
};

const EDITOR_TRANSLATION_KEY: InjectionKey<{
  locale: MaybeRefOrGetter<string>;
  translation: ReturnType<typeof useLocaleTranslator>;
}> = Symbol("folioEditorTranslation");

const translateWith =
  (
    translator: ReturnType<typeof useLocaleTranslator>["translator"],
  ): TranslateFn =>
  (key, values) =>
    // SAFETY: use-intl's translator accepts the paths and ICU values used by SFCs.
    (translator.value as TranslateFn)(key, values);

export const useEditorTranslation = (onError: (error: Error) => void) => {
  const locale = inject(LOCALE_KEY, defaultLocale);
  const translation = useLocaleTranslator(locale, onError);
  provide(EDITOR_TRANSLATION_KEY, { locale, translation });
  return {
    t: translateWith(translation.translator),
    status: translation.status,
  };
};

/**
 * Access folio's translator. Returns `{ t }` where `t("key")` looks up the
 * `folio` namespace of the active locale's catalog.
 */
export const useTranslation = (): { t: TranslateFn } => {
  const locale = inject(LOCALE_KEY, defaultLocale);
  const inherited = inject(EDITOR_TRANSLATION_KEY, null);
  const { translator } =
    inherited?.locale === locale
      ? inherited.translation
      : useLocaleTranslator(locale);
  return { t: translateWith(translator) };
};

/**
 * Vue plugin form: `app.use(i18nPlugin, "de")` provides the translator at the
 * app root so every folio component resolves translations without a wrapping
 * provider component.
 */
export const i18nPlugin = {
  install(app: App, locale: MaybeRefOrGetter<string> = defaultLocale): void {
    app.provide(LOCALE_KEY, locale);
  },
};
