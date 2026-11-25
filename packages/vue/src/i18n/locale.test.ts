import { expect, test } from "bun:test";
import { createSSRApp, h } from "vue";
import { renderToString } from "vue/server-renderer";

import { i18nPlugin, provideLocale, useTranslation } from ".";

const Label = {
  setup() {
    const { t } = useTranslation();
    return () => h("button", t("bold"));
  },
};

test("the global plugin and a nested provider await their locale during SSR", async () => {
  const Nested = {
    setup() {
      provideLocale("pt-BR");
      return () => h(Label);
    },
  };
  const app = createSSRApp({
    render: () => h("div", [h(Label), h(Nested)]),
  });
  app.use(i18nPlugin, "de-DE");
  expect(await renderToString(app)).toBe(
    "<div><button>Fett</button><button>Negrito</button></div>",
  );
});
