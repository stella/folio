import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act } from "react";
import { createRoot } from "react-dom/client";

import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";

import { FormattingBar } from "./FormattingBar";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

// The toolbar's own named button is the color picker's trigger; the picker
// must not wrap it in a second, unnamed element with button semantics.
test("the font color button is the picker trigger itself", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () =>
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <FormattingBar
            enableShortcuts={false}
            showStylePicker={false}
            showFontPicker={false}
            showFontSizePicker={false}
            showAlignmentButtons={false}
            showListButtons={false}
          />
        </IntlProvider>,
      ),
    );
    const trigger = container.querySelector<HTMLButtonElement>(
      'button[aria-label="Font color"]',
    );
    if (!trigger) throw new Error("Font color button did not mount");

    expect(trigger.parentElement?.closest("button, [role='button']")).toBe(
      null,
    );
    expect(trigger.querySelector("button, [role='button']")).toBe(null);
    expect(trigger.getAttribute("aria-expanded")).toBe("false");

    await act(async () => trigger.click());

    expect(trigger.getAttribute("aria-haspopup")).toBe("dialog");
    expect(trigger.getAttribute("aria-expanded")).toBe("true");
    expect(trigger.getAttribute("aria-controls")).not.toBe(null);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});
