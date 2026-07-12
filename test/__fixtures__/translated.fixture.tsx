// Clean fixture for scripts/no-untranslated-jsx-literal.test.ts.
// Every user-facing string goes through use-intl, so the
// no-untranslated-jsx-literal rule must report nothing here.
import { useTranslations } from "use-intl";

export const TranslatedFixture = () => {
  const t = useTranslations("folio");
  return <button type="button">{t("common.apply")}</button>;
};
