import { createApp } from "vue";

import App from "./App.vue";

const container = document.querySelector("#app");
if (container) {
  createApp(App).mount(container);
}
