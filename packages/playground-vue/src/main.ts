import { createApp } from "vue";

import {
  type ColorMode,
  colorModePlugin,
  i18nPlugin,
  renderAsync,
} from "@stll/folio-vue";
import "@stll/folio-vue/editor.css";

import App from "./App.vue";

const COLOR_MODES = [
  "light",
  "dark",
  "system",
] as const satisfies readonly ColorMode[];

function isColorMode(value: string | null): value is ColorMode {
  return COLOR_MODES.some((mode) => mode === value);
}

const container = document.querySelector<HTMLElement>("#app");
const search = new URLSearchParams(window.location.search);
const mountLocale = search.get("mountLocale");
if (container && mountLocale !== null) {
  void (async () => {
    const response = await fetch("/fixtures/sample.docx");
    await renderAsync(await response.arrayBuffer(), container, {
      locale: mountLocale,
    });
  })();
} else if (container) {
  const requestedColorMode = search.get("colorMode");
  const app = createApp(App);
  app.use(i18nPlugin, search.get("locale") ?? "en");
  app.use(
    colorModePlugin,
    isColorMode(requestedColorMode) ? requestedColorMode : "light",
  );
  app.mount(container);
}
