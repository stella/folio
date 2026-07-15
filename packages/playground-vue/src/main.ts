import { createApp } from "vue";
import { colorModePlugin, i18nPlugin, type ColorMode } from "@stll/folio-vue";

import App from "./App.vue";

const COLOR_MODES = ["light", "dark", "system"] as const satisfies readonly ColorMode[];

function isColorMode(value: string | null): value is ColorMode {
  return COLOR_MODES.some((mode) => mode === value);
}

const container = document.querySelector("#app");
if (container) {
  const search = new URLSearchParams(window.location.search);
  const requestedColorMode = search.get("colorMode");
  const app = createApp(App);
  app.use(i18nPlugin, search.get("locale") ?? "en");
  app.use(colorModePlugin, isColorMode(requestedColorMode) ? requestedColorMode : "light");
  app.mount(container);
}
