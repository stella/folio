import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act } from "react";
import { createRoot } from "react-dom/client";

import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { PAINTER_PAINTED_EVENT } from "@stll/folio-core/layout-painter/renderPage";

import { InlineHeaderFooterEditor } from "./InlineHeaderFooterEditor";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});
const noop = () => {};
const noView = () => null;

/** A painted header slot that measures at `top`. */
const slotAt = (top: number): HTMLElement => {
  const slot = document.createElement("div");
  slot.className = "layout-page-header";
  slot.getBoundingClientRect = () => new DOMRect(40, top, 500, 30);
  return slot;
};

test("the chrome follows the slot a repaint puts in place of the old one", async () => {
  const parentElement = document.createElement("div");
  parentElement.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
  const pages = document.createElement("div");
  parentElement.append(pages);
  document.body.append(parentElement);
  pages.append(slotAt(100));
  const host = document.createElement("div");
  document.body.append(host);
  const root = createRoot(host);
  const chromeTop = () =>
    host.querySelector<HTMLElement>("[style*='position: absolute']")?.style.top;
  try {
    // The host renders once; repaints that follow commit nothing in it.
    await act(async () =>
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <InlineHeaderFooterEditor
            position="header"
            targetSelector=".layout-page-header"
            parentElement={parentElement}
            getActiveView={noView}
            onClose={noop}
          />
        </IntlProvider>,
      ),
    );
    expect(chromeTop()).toBe("100px");

    // The painter replaces the slot when it repaints the page.
    pages.replaceChildren(slotAt(260));
    await act(async () => {
      pages.dispatchEvent(
        new CustomEvent(PAINTER_PAINTED_EVENT, { bubbles: true }),
      );
    });
    expect(chromeTop()).toBe("260px");
  } finally {
    await act(async () => root.unmount());
    parentElement.remove();
    host.remove();
  }
});
