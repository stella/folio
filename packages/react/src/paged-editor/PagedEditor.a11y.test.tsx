import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act } from "react";
import { createRoot } from "react-dom/client";

import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";

import { PagedEditor } from "./PagedEditor";

const documentIO = {
  getDocx: async () => new ArrayBuffer(0),
  loadDocument: () => undefined,
  loadDocx: async () => undefined,
};

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

test.each([
  { readOnly: true, role: "document", multiline: null },
  { readOnly: false, role: "textbox", multiline: "true" },
])(
  "editor root exposes named $role semantics",
  async ({ readOnly, role, multiline }) => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const accessibleName = "Employment agreement";

    try {
      await act(async () => {
        root.render(
          <IntlProvider
            locale="en"
            timeZone="UTC"
            messages={getFolioMessages("en")}
          >
            <PagedEditor
              accessibleName={accessibleName}
              document={null}
              documentIdentity="accessible-document"
              documentIO={documentIO}
              markupView="all-markup"
              readOnly={readOnly}
            />
          </IntlProvider>,
        );
      });

      const editorRoot = container.querySelector(".paged-editor");
      expect(editorRoot).not.toBeNull();
      expect(editorRoot?.getAttribute("role")).toBe(role);
      expect(editorRoot?.getAttribute("aria-label")).toBe(accessibleName);
      expect(editorRoot?.getAttribute("aria-multiline")).toBe(multiline);
      expect(editorRoot?.getAttribute("tabindex")).toBe("0");
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  },
);
