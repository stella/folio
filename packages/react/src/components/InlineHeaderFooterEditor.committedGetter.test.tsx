import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, mock, test } from "bun:test";
import { act, startTransition, Suspense } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { InlineHeaderFooterEditor } from "./InlineHeaderFooterEditor";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});
const pending = new Promise<never>(() => {});
const noop = () => {};
const Suspend = () => {
  throw pending;
};

test("the focus loop reads only the latest committed active-view getter", async () => {
  const originalRequest = globalThis.requestAnimationFrame;
  const originalCancel = globalThis.cancelAnimationFrame;
  const frames = new Map<number, FrameRequestCallback>();
  let nextId = 0;
  globalThis.requestAnimationFrame = (callback) => {
    const id = ++nextId;
    frames.set(id, callback);
    return id;
  };
  globalThis.cancelAnimationFrame = (id) => {
    frames.delete(id);
  };
  const advanceFrame = async () => {
    const callbacks = [...frames.values()];
    frames.clear();
    await act(async () => {
      for (const callback of callbacks) callback(0);
    });
  };
  const oldGetter = mock(() => null);
  const uncommittedGetter = mock(() => null);
  const newGetter = mock(() => null);
  const parentElement = document.createElement("div");
  const targetElement = document.createElement("div");
  parentElement.append(targetElement);
  const root = createRoot(document.createElement("div"));
  const render = (getter: () => null, suspend = false) => (
    <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
      <Suspense fallback={null}>
        <InlineHeaderFooterEditor
          position="header"
          targetElement={targetElement}
          parentElement={parentElement}
          getActiveView={getter}
          onClose={noop}
        />
        {suspend && <Suspend />}
      </Suspense>
    </IntlProvider>
  );
  try {
    await act(async () => root.render(render(oldGetter)));
    await act(async () => {
      startTransition(() => root.render(render(uncommittedGetter, true)));
    });
    await advanceFrame();
    expect(oldGetter).toHaveBeenCalled();
    expect(uncommittedGetter).not.toHaveBeenCalled();

    oldGetter.mockClear();
    await act(async () => root.render(render(newGetter)));
    await advanceFrame();
    expect(newGetter).toHaveBeenCalled();
    expect(oldGetter).not.toHaveBeenCalled();
  } finally {
    await act(async () => root.unmount());
    globalThis.requestAnimationFrame = originalRequest;
    globalThis.cancelAnimationFrame = originalCancel;
  }
});
