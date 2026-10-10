import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, useLayoutEffect } from "react";
import { createRoot } from "react-dom/client";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";
import { usePanelLayout } from "./usePanelLayout";
import { useActiveHeading } from "./useActiveHeading";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createResizeFixture = () => {
  const observers: ObservedResize[] = [];
  class ObservedResize implements ResizeObserver {
    readonly targets = new Set<Element>();
    constructor(private readonly callback: ResizeObserverCallback) {
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
  return {
    Observer: ObservedResize,
    notify: (target: Element) => {
      for (const observer of observers) observer.notify(target);
    },
    isObserved: (target: Element) => observers.some((observer) => observer.targets.has(target)),
  };
};

const createScrollWidthFixture = (initialScrollbar: number) => {
  const element = document.createElement("div");
  let scrollbar = initialScrollbar;
  Object.defineProperties(element, {
    offsetWidth: { configurable: true, get: () => 600 + scrollbar },
    clientWidth: { configurable: true, get: () => 600 },
  });
  return {
    element,
    setScrollbar: (width: number) => {
      scrollbar = width;
    },
  };
};

function PanelProbe({
  scrollContainer,
  row,
}: {
  scrollContainer: HTMLElement | null;
  row: HTMLDivElement;
}) {
  const { rowRef, layout } = usePanelLayout({
    scrollContainer,
    pageWidth: 600,
    outline: "available",
    comments: "closed",
  });
  useLayoutEffect(() => {
    rowRef(row);
    return () => rowRef(null);
  }, [row, rowRef]);
  return <output data-outline={layout.outline} />;
}

// Fixtures previously retained one DOM root. Replace it while keeping the row
// and hook instance alive, and emit resize only from actually observed roots.
test("panel measurements rebind when the scroll container is replaced", async () => {
  const previousObserver = globalThis.ResizeObserver;
  const resize = createResizeFixture();
  globalThis.ResizeObserver = resize.Observer;
  const host = document.createElement("div");
  const row = document.createElement("div");
  Object.defineProperty(row, "clientWidth", { configurable: true, value: 900 });
  document.body.append(host, row);
  const root = createRoot(host);
  const first = createScrollWidthFixture(0);
  const replacement = createScrollWidthFixture(20);
  try {
    await act(async () => root.render(<PanelProbe scrollContainer={first.element} row={row} />));
    expect(host.querySelector("output")?.dataset["outline"]).toBe("column");
    expect(resize.isObserved(first.element)).toBe(true);
    await act(async () =>
      root.render(<PanelProbe scrollContainer={replacement.element} row={row} />),
    );
    expect(resize.isObserved(first.element)).toBe(false);
    expect(resize.isObserved(replacement.element)).toBe(true);
    expect(resize.isObserved(row)).toBe(true);
    expect(host.querySelector("output")?.dataset["outline"]).toBe("rail");

    first.setScrollbar(300);
    await act(async () => resize.notify(first.element));
    expect(host.querySelector("output")?.dataset["outline"]).toBe("rail");
    replacement.setScrollbar(0);
    await act(async () => resize.notify(replacement.element));
    expect(host.querySelector("output")?.dataset["outline"]).toBe("column");
    await act(async () => root.render(<PanelProbe scrollContainer={null} row={row} />));
    expect(resize.isObserved(replacement.element)).toBe(false);
    expect(resize.isObserved(row)).toBe(true);
  } finally {
    await act(async () => root.unmount());
    expect(resize.isObserved(row)).toBe(false);
    globalThis.ResizeObserver = previousObserver;
    host.remove();
    row.remove();
  }
});

const headings = [
  { text: "First", level: 0, pmPos: 0 },
  { text: "Second", level: 1, pmPos: 7 },
] satisfies HeadingInfo[];

const createHeadingRoot = (initialSecondTop: number) => {
  const element = document.createElement("div");
  element.innerHTML =
    '<div class="layout-page"><div class="layout-page-content"><p><span data-pm-start="1" data-pm-end="6">First</span></p><p><span data-pm-start="8" data-pm-end="14">Second</span></p></div></div>';
  Object.defineProperty(element, "clientHeight", { configurable: true, value: 200 });
  element.getBoundingClientRect = () => new DOMRect(0, 0, 600, 200);
  const first = element.querySelector('[data-pm-start="1"]');
  const second = element.querySelector('[data-pm-start="8"]');
  if (!(first instanceof HTMLElement) || !(second instanceof HTMLElement)) {
    panic("Heading fixture must contain its painted body spans");
  }
  let secondTop = initialSecondTop;
  first.getBoundingClientRect = () => new DOMRect(0, 20, 100, 16);
  second.getBoundingClientRect = () => new DOMRect(0, secondTop, 100, 16);
  return {
    element,
    setSecondTop: (top: number) => {
      secondTop = top;
    },
  };
};

const createFrameFixture = () => {
  const frames = new Map<number, FrameRequestCallback>();
  let nextId = 0;
  return {
    request: (callback: FrameRequestCallback) => {
      nextId++;
      frames.set(nextId, callback);
      return nextId;
    },
    cancel: (id: number) => {
      frames.delete(id);
    },
    flush: () => {
      const pending = [...frames.values()];
      frames.clear();
      for (const callback of pending) callback(0);
    },
    pending: () => frames.size,
  };
};

function HeadingProbe({ scrollContainer }: { scrollContainer: HTMLElement | null }) {
  const { activeId } = useActiveHeading(scrollContainer, headings);
  return <output data-active={activeId ?? "none"} />;
}

test("active heading follows only the replacement scroll root and cancels old work", async () => {
  const previousRequest = globalThis.requestAnimationFrame;
  const previousCancel = globalThis.cancelAnimationFrame;
  const frames = createFrameFixture();
  globalThis.requestAnimationFrame = frames.request;
  globalThis.cancelAnimationFrame = frames.cancel;
  const host = document.createElement("div");
  const first = createHeadingRoot(200);
  const replacement = createHeadingRoot(80);
  document.body.append(host, first.element, replacement.element);
  const root = createRoot(host);
  try {
    await act(async () => root.render(<HeadingProbe scrollContainer={first.element} />));
    await act(async () => frames.flush());
    expect(host.querySelector("output")?.dataset["active"]).toBe("0");
    first.element.dispatchEvent(new Event("scroll"));
    expect(frames.pending()).toBe(1);
    await act(async () => root.render(<HeadingProbe scrollContainer={replacement.element} />));
    expect(frames.pending()).toBe(1);
    await act(async () => frames.flush());
    expect(host.querySelector("output")?.dataset["active"]).toBe("7");

    first.setSecondTop(80);
    first.element.dispatchEvent(new Event("scroll"));
    expect(frames.pending()).toBe(0);
    expect(host.querySelector("output")?.dataset["active"]).toBe("7");
    replacement.setSecondTop(200);
    replacement.element.dispatchEvent(new Event("scroll"));
    expect(frames.pending()).toBe(1);
    await act(async () => frames.flush());
    expect(host.querySelector("output")?.dataset["active"]).toBe("0");

    replacement.element.dispatchEvent(new Event("scroll"));
    await act(async () => root.render(<HeadingProbe scrollContainer={null} />));
    expect(frames.pending()).toBe(0);
    replacement.element.dispatchEvent(new Event("scroll"));
    expect(frames.pending()).toBe(0);
  } finally {
    await act(async () => root.unmount());
    globalThis.requestAnimationFrame = previousRequest;
    globalThis.cancelAnimationFrame = previousCancel;
    host.remove();
    first.element.remove();
    replacement.element.remove();
  }
});
