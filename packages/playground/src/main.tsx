import "./styles.css";
import "@stll/folio-react/editor.css";

import { createRoot } from "react-dom/client";

import { renderAsync } from "@stll/folio-react";

import { App } from "./App";

const container = document.querySelector<HTMLElement>("#app");
const mountLocale = new URLSearchParams(window.location.search).get(
  "mountLocale",
);
if (container && mountLocale !== null) {
  void (async () => {
    const response = await fetch("/fixtures/sample.docx");
    await renderAsync(await response.arrayBuffer(), container, {
      locale: mountLocale,
      accessibleName: "Locale test editor",
    });
  })();
} else if (container) {
  const root = createRoot(container);
  root.render(<App />);
}
