import { afterAll, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { resetFolioMessageAssets } from "@stll/folio-core/i18n/loadMessages";
import {
  configureFolioMessageAssets,
  loadFolioMessages,
} from "@stll/folio-core/i18n/messages";

import { acquireHarnessDom } from "../../../../test/harnessDom";

const releaseDom = acquireHarnessDom();
afterAll(releaseDom);

const { createApp, createSSRApp, h, nextTick, ref } = await import("vue");
const { renderToString } = await import("vue/server-renderer");
const { i18nPlugin, provideLocale, useEditorTranslation, useTranslation } =
  await import(".");

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

test("a locale switch hides translated UI until its catalog arrives", async () => {
  resetFolioMessageAssets();
  configureFolioMessageAssets({ baseUrl: new URL("https://catalogs.test/") });
  const originalFetch = globalThis.fetch;
  let release: ((response: Response) => void) | undefined;
  globalThis.fetch = () =>
    new Promise<Response>((resolve) => {
      release = resolve;
    });
  const locale = ref("en");
  const errors: Error[] = [];
  const EditorLabel = {
    setup() {
      const { status, t } = useEditorTranslation((error) => errors.push(error));
      return () =>
        status.value.type === "loading"
          ? h("span", "loading")
          : h("button", t("bold"));
    },
  };
  const app = createApp(EditorLabel);
  app.use(i18nPlugin, locale);
  const container = document.createElement("div");
  app.mount(container);
  try {
    expect(container.textContent).toBe("Bold");
    locale.value = "cs";
    await nextTick();
    expect(container.textContent).toBe("loading");
    const czech = readFileSync(
      path.join(import.meta.dir, "../../../core/src/i18n/messages/cs.json"),
      "utf8",
    );
    release?.(new Response(czech));
    await loadFolioMessages("cs");
    await nextTick();
    expect(container.textContent).toBe("Tučné");
    expect(errors).toEqual([]);
  } finally {
    app.unmount();
    globalThis.fetch = originalFetch;
    resetFolioMessageAssets();
  }
});

test("a failed locale switch reports the error and paints English", async () => {
  resetFolioMessageAssets();
  configureFolioMessageAssets({ baseUrl: new URL("https://catalogs.test/") });
  const originalFetch = globalThis.fetch;
  globalThis.fetch = () => Promise.resolve(new Response(null, { status: 503 }));
  const locale = ref("en");
  const errors: Error[] = [];
  const EditorLabel = {
    setup() {
      const { status, t } = useEditorTranslation((error) => errors.push(error));
      return () =>
        status.value.type === "loading"
          ? h("span", "loading")
          : h("button", t("bold"));
    },
  };
  const app = createApp(EditorLabel);
  app.use(i18nPlugin, locale);
  const container = document.createElement("div");
  app.mount(container);
  try {
    locale.value = "cs";
    await nextTick();
    expect(container.textContent).toBe("loading");
    await new Promise((resolve) => setTimeout(resolve, 0));
    await nextTick();
    expect(container.textContent).toBe("Bold");
    expect(errors).toHaveLength(1);
  } finally {
    app.unmount();
    globalThis.fetch = originalFetch;
    resetFolioMessageAssets();
  }
});
