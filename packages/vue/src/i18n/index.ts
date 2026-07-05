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
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { createTranslator } from "use-intl/core";
import {
  computed,
  inject,
  provide,
  toValue,
  type App,
  type ComputedRef,
  type InjectionKey,
  type MaybeRefOrGetter,
} from "vue";

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
export type TranslateFn = (key: string, values?: Record<string, unknown>) => string;

const createFolioTranslator = (locale: string) =>
  createTranslator({ locale, messages: getFolioMessages(locale), namespace: "folio" });

const TRANSLATOR_KEY: InjectionKey<ComputedRef<FolioTranslator>> = Symbol("folioTranslator");

// Standalone fallback: when no provider is mounted, default to English so an
// un-provided `<DocxEditor>` still renders localized chrome instead of throwing.
const fallbackTranslator = computed(() => createFolioTranslator(defaultLocale));

/**
 * Provide folio's translator to descendant components. `locale` may be a ref,
 * getter, or plain string; the translator recomputes when a reactive locale
 * changes. Call inside a parent component `setup()` (e.g. the editor shell).
 */
export const provideLocale = (locale: MaybeRefOrGetter<string> = defaultLocale): void => {
  provide(
    TRANSLATOR_KEY,
    computed(() => createFolioTranslator(toValue(locale))),
  );
};

/**
 * Access folio's translator. Returns `{ t }` where `t("key")` looks up the
 * `folio` namespace of the active locale's catalog.
 */
export const useTranslation = (): { t: TranslateFn } => {
  const translator = inject(TRANSLATOR_KEY, fallbackTranslator);
  // Unwrap the computed into a stable callable so call sites use `t("key")`
  // without touching `.value`.
  const t: TranslateFn = (key, values) =>
    // SAFETY: use-intl's translator is callable with (key, values); the cast
    // narrows the broad overload set to the shape folio's SFCs use.
    (translator.value as (k: string, v?: Record<string, unknown>) => string)(key, values);
  return { t };
};

/**
 * Vue plugin form: `app.use(i18nPlugin, "de")` provides the translator at the
 * app root so every folio component resolves translations without a wrapping
 * provider component.
 */
export const i18nPlugin = {
  install(app: App, locale: MaybeRefOrGetter<string> = defaultLocale): void {
    app.provide(
      TRANSLATOR_KEY,
      computed(() => createFolioTranslator(toValue(locale))),
    );
  },
};
