import "./styles.css";
import "@stll/folio-react/editor.css";

import { createRoot } from "react-dom/client";

import { App } from "./App";

const container = document.querySelector("#app");
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}
