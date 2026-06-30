import "./styles.css";
import "@stll/folio-react/editor.css";

import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { App } from "./App";
import messages from "./messages/en.json";

const container = document.querySelector("#app");
if (container) {
  const root = createRoot(container);
  root.render(
    <IntlProvider
      locale="en"
      messages={messages}
      timeZone={Intl.DateTimeFormat().resolvedOptions().timeZone}
    >
      <App />
    </IntlProvider>,
  );
}
