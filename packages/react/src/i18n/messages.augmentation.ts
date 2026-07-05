import type FolioMessages from "@stll/folio-core/i18n/messages/messages.gen";

// Type folio's own UI translations against the generated English catalog. With
// this augmentation `useTranslations("folio")` and every `t("key")` call in
// folio's components are checked against `messages.gen.ts` (regenerated from
// en.json by `scripts/i18n-typegen.ts`): a key that the catalog does not carry
// is a typecheck error, not a runtime `MISSING_MESSAGE`. Ported from
// stella/stella's `apps/web/src/types/i18n.d.ts`.
//
// This is a regular `.ts` module, not a `.d.ts`: the base tsconfig sets
// `skipLibCheck`, which skips declaration-file bodies, so an augmentation placed
// in a `.d.ts` never applies. As a plain source module it is type-checked and
// the augmentation takes effect.
//
// It stays internal to folio's own typecheck: no shipped entry (`index.ts`,
// `i18n/messages.ts`) imports it, so the package build never emits it into dist
// and it does not leak `AppConfig.Messages` onto consumers, who declare their
// own catalog when they merge in `getFolioMessages(locale)`.
declare module "use-intl" {
  // oxlint-disable-next-line typescript/consistent-type-definitions -- module augmentation requires interface merging
  interface AppConfig {
    Messages: FolioMessages;
  }
}
