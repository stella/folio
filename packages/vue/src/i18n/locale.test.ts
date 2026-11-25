import { afterAll, expect, test } from "bun:test";

import { loadFolioMessages } from "@stll/folio-core/i18n/loadMessages";

import { acquireHarnessDom } from "../../../../test/harnessDom";

const releaseDom = acquireHarnessDom();
afterAll(releaseDom);

const { createApp, createSSRApp, h, nextTick, ref } = await import("vue");
const { renderToString } = await import("vue/server-renderer");
const { i18nPlugin, provideLocale, useTranslation } = await import(".");

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

test("a late catalog cannot replace a newer locale", async () => {
  const locale = ref("en");
  const app = createApp(Label);
  app.use(i18nPlugin, locale);
  const container = document.createElement("div");
  app.mount(container);
  try {
    expect(container.textContent).toBe("Bold");
    locale.value = "cs";
    locale.value = "en";
    await loadFolioMessages("cs");
    await nextTick();
    expect(container.textContent).toBe("Bold");

    locale.value = "cs";
    await nextTick();
    expect(container.textContent).toBe("Tučné");
  } finally {
    app.unmount();
  }
});
