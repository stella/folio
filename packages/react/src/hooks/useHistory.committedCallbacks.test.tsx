import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, mock, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef, Suspense, useImperativeHandle } from "react";
import { createRoot } from "react-dom/client";
import { useHistory, type UseHistoryReturn, type UseHistoryOptions } from "./useHistory";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});
const pending = new Promise<never>(() => {});

const oldUndo = mock((_state: string) => {});
const newUndo = mock((_state: string) => {});
const newRedo = mock((_state: string) => {});
const oldOptions = { onUndo: oldUndo };
const suspendedOptions = { onUndo: newUndo };
const committedOptions = { onUndo: newUndo, onRedo: newRedo };

test("stable history commands use the latest committed callbacks and reset value", async () => {
  const ref = createRef<UseHistoryReturn<string>>();
  const Harness = ({
    initial,
    options,
    suspend = false,
  }: {
    initial: string;
    options: UseHistoryOptions<string>;
    suspend?: boolean;
  }) => {
    const history = useHistory(initial, {
      ...options,
      enableKeyboardShortcuts: false,
      groupingInterval: 0,
    });
    useImperativeHandle(ref, () => history, [history]);
    if (suspend) throw pending;
    return <output>{history.state}</output>;
  };
  const container = document.createElement("div");
  const root = createRoot(container);
  try {
    await act(async () =>
      root.render(
        <Suspense fallback={null}>
          <Harness initial="first" options={oldOptions} />
        </Suspense>,
      ),
    );
    const api = ref.current ?? panic("History handle missing");
    await act(async () => api.push("edited"));
    await act(async () =>
      root.render(
        <Suspense fallback={null}>
          <Harness initial="uncommitted" options={suspendedOptions} suspend />
        </Suspense>,
      ),
    );
    await act(async () => {
      api.undo();
    });
    expect(oldUndo).toHaveBeenCalledWith("first");
    expect(newUndo).not.toHaveBeenCalled();
    await act(async () =>
      root.render(
        <Suspense fallback={null}>
          <Harness initial="committed" options={committedOptions} />
        </Suspense>,
      ),
    );
    expect(ref.current?.undo).toBe(api.undo);
    expect(ref.current?.redo).toBe(api.redo);
    expect(ref.current?.reset).toBe(api.reset);
    await act(async () => {
      api.redo();
    });
    expect(newRedo).toHaveBeenCalledWith("edited");
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      api.reset();
    });
    expect(container.textContent).toBe("committed");
    await act(async () => {
      api.push("second edit");
      api.undo();
    });
    expect(newUndo).toHaveBeenCalledWith("committed");
  } finally {
    await act(async () => root.unmount());
  }
});
