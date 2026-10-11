import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { act } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";

import { FormattingBar } from "./FormattingBar";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

// Resize events must come from observed elements: the former fixture never
// detached/remounted the secondary group before changing its natural width.
test("secondary controls remain observed after collapsing and remounting", async () => {
  const previousObserver = globalThis.ResizeObserver;
  const previousClientWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, "clientWidth");
  const previousOffsetWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, "offsetWidth");
  const observers: ObservedResize[] = [];
  class ObservedResize implements ResizeObserver {
    readonly targets = new Set<Element>();
    private readonly callback: ResizeObserverCallback;
    constructor(callback: ResizeObserverCallback) {
      this.callback = callback;
      observers.push(this);
    }
    observe(target: Element) {
      this.targets.add(target);
    }
    unobserve(target: Element) {
      this.targets.delete(target);
    }
    disconnect() {
      this.targets.clear();
    }
    notify(target: Element) {
      if (this.targets.has(target)) this.callback([], this);
    }
  }
  globalThis.ResizeObserver = ObservedResize;
  let available = 800;
  let secondaryWidth = 300;
  Object.defineProperty(HTMLElement.prototype, "clientWidth", {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains("overflow-x-auto") ? available : 0;
    },
  });
  Object.defineProperty(HTMLElement.prototype, "offsetWidth", {
    configurable: true,
    get(this: HTMLElement) {
      if (!this.parentElement?.classList.contains("overflow-x-auto")) return 0;
      return this.querySelector("[data-secondary-probe]") ? secondaryWidth : 200;
    },
  });
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const toolbarProps = {
    enableShortcuts: false,
    showStylePicker: false,
    showFontPicker: false,
    showFontSizePicker: false,
    showTextColorPicker: false,
    showAlignmentButtons: false,
    showListButtons: false,
    inlineExtra: <span data-secondary-probe="true" />,
  };
  try {
    await act(async () =>
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <FormattingBar {...toolbarProps} />
        </IntlProvider>,
      ),
    );
    const scroll = container.querySelector(".overflow-x-auto");
    if (!scroll) throw new Error("Toolbar scroll region did not mount");
    const secondary = () =>
      scroll.querySelector("[data-secondary-probe]")?.parentElement?.parentElement;
    const initial = secondary();
    expect(initial).toBeDefined();
    const notify = async (target: Element) => {
      await act(async () => {
        for (const observer of observers) observer.notify(target);
      });
    };
    available = 350;
    await notify(scroll);
    expect(secondary()).toBeUndefined();
    expect(observers.some((observer) => initial && observer.targets.has(initial))).toBe(false);
    available = 800;
    await notify(scroll);
    const remounted = secondary();
    if (!remounted) throw new Error("Secondary group did not remount");
    expect(remounted).not.toBe(initial);
    expect(observers.some((observer) => observer.targets.has(remounted))).toBe(true);
    secondaryWidth = 700;
    await notify(remounted);
    expect(secondary()).toBeUndefined();
  } finally {
    await act(async () => root.unmount());
    container.remove();
    globalThis.ResizeObserver = previousObserver;
    if (previousClientWidth)
      Object.defineProperty(HTMLElement.prototype, "clientWidth", previousClientWidth);
    else Reflect.deleteProperty(HTMLElement.prototype, "clientWidth");
    if (previousOffsetWidth)
      Object.defineProperty(HTMLElement.prototype, "offsetWidth", previousOffsetWidth);
    else Reflect.deleteProperty(HTMLElement.prototype, "offsetWidth");
  }
});
