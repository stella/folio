import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import {
  afterAll,
  afterEach,
  beforeEach,
  describe,
  expect,
  test,
} from "bun:test";

import type {
  MeasuredLine,
  Page,
  ParagraphBlock,
  ParagraphFragment,
  ParagraphMeasure,
} from "../layout-engine/types";
import type { BlockLookup } from "../layout-painter/index";
import { PAGE_CLASS_NAMES, renderPages } from "../layout-painter/renderPage";
import { scrollPagesToPmPosition } from "./scrollToPmPosition";

afterAll(() => GlobalRegistrator.unregister());

const PAGE_COUNT = 78;
/** Each page holds one paragraph spanning `[index * 100, (index + 1) * 100)`. */
const PAGE_PM_SPAN = 100;
const PAGE_HEIGHT = 1056;
const PAGE_PITCH = 1080;
const VIEWPORT_HEIGHT = 900;
/** Where a painted element sits below the top of its page. */
const CONTENT_OFFSET = 100;
const CONTENT_HEIGHT = 12;
/** Frames a smooth scroll takes to arrive, whatever the distance. */
const SMOOTH_SCROLL_FRAMES = 30;
const MAX_FRAMES = 200;

const pageStart = (index: number) => index * PAGE_PM_SPAN;

const buildPages = (): { pages: Page[]; blockLookup: BlockLookup } => {
  const pages: Page[] = [];
  const blockLookup: BlockLookup = new Map();
  for (let index = 0; index < PAGE_COUNT; index++) {
    const text = `Clause ${String(index + 1)}`;
    const id = `clause-${String(index)}`;
    const block: ParagraphBlock = {
      kind: "paragraph",
      id,
      runs: [
        {
          kind: "text",
          text,
          pmStart: pageStart(index) + 1,
          pmEnd: pageStart(index) + 1 + text.length,
        },
      ],
      pmStart: pageStart(index),
      pmEnd: pageStart(index + 1),
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: text.length,
      width: 80,
      ascent: 8,
      descent: 2,
      lineHeight: CONTENT_HEIGHT,
    };
    const measure: ParagraphMeasure = {
      kind: "paragraph",
      lines: [line],
      totalHeight: CONTENT_HEIGHT,
    };
    const fragment: ParagraphFragment = {
      kind: "paragraph",
      blockId: id,
      x: 72,
      y: 72,
      width: 600,
      height: CONTENT_HEIGHT,
      fromLine: 0,
      toLine: 1,
      pmStart: pageStart(index),
      pmEnd: pageStart(index + 1),
    };
    blockLookup.set(id, { block, measure });
    pages.push({
      number: index + 1,
      logicalNumber: index + 1,
      sectionPageNumber: index + 1,
      fragments: [fragment],
      margins: { top: 72, right: 72, bottom: 72, left: 72 },
      size: { w: 816, h: PAGE_HEIGHT },
    });
  }
  return { pages, blockLookup };
};

type ObserverCallback = (entries: IntersectionObserverEntry[]) => void;

/**
 * An observer the viewport below drives frame by frame, the way a browser
 * delivers entries: once per observed target, then on every change.
 */
class FakeIntersectionObserver {
  static instances: FakeIntersectionObserver[] = [];
  readonly callback: ObserverCallback;
  /** Extra band above and below the viewport, in pixels. */
  readonly margin: number;
  readonly targets = new Map<Element, boolean | null>();

  constructor(callback: ObserverCallback, options?: IntersectionObserverInit) {
    this.callback = callback;
    const top = (options?.rootMargin ?? "0px").split(" ").at(0) ?? "0px";
    this.margin = top.endsWith("%")
      ? (Number.parseFloat(top) / 100) * VIEWPORT_HEIGHT
      : Number.parseFloat(top);
    FakeIntersectionObserver.instances.push(this);
  }

  observe(target: Element): void {
    this.targets.set(target, null);
  }
  unobserve(target: Element): void {
    this.targets.delete(target);
  }
  disconnect(): void {
    this.targets.clear();
    FakeIntersectionObserver.instances =
      FakeIntersectionObserver.instances.filter(
        (observer) => observer !== this,
      );
  }
}

type ScrollCall = { top: number; behavior: ScrollBehavior | undefined };

type JumpTrace = {
  /** Pages painted by the navigation call itself, before any frame. */
  paintedBeforeScroll: number[];
  /** Pages painted between the call and the viewport settling. */
  paintedPages: number[];
  /** Frames whose viewport centre showed a page without content. */
  blankFrames: number;
  /** Page measurements the painter made while the observers ran. */
  pageMeasurements: number;
  scrollCalls: ScrollCall[];
};

const originalObserver = globalThis.IntersectionObserver;
const originalFrame = globalThis.requestAnimationFrame;
const originalRect = Object.getOwnPropertyDescriptor(
  HTMLElement.prototype,
  "getBoundingClientRect",
);
const originalInnerHeight = Object.getOwnPropertyDescriptor(
  window,
  "innerHeight",
);

/**
 * Pages stacked at a fixed pitch inside the editor's scroll root, which fills
 * the viewport. Smooth scrolls animate over a fixed number of frames; each
 * frame runs queued animation callbacks, then delivers intersection changes.
 */
const mountPages = () => {
  const root = document.createElement("div");
  root.setAttribute("data-folio-scroll", "");
  const container = document.createElement("div");
  root.append(container);
  document.body.append(root);
  let scrollTop = 0;
  let animation: { from: number; to: number; frame: number } | null = null;
  let measuring = false;
  let pageMeasurements = 0;
  const scrollCalls: ScrollCall[] = [];
  const frameCallbacks: FrameRequestCallback[] = [];
  Object.defineProperties(root, {
    clientHeight: { value: VIEWPORT_HEIGHT, configurable: true },
    clientWidth: { value: 1000, configurable: true },
    scrollTop: {
      configurable: true,
      get: () => scrollTop,
      set: (value: number) => {
        scrollTop = value;
      },
    },
  });
  root.scrollTo = (options?: ScrollToOptions | number) => {
    if (typeof options !== "object") {
      throw new TypeError("The editor scrolls with options");
    }
    const top = options.top ?? scrollTop;
    scrollCalls.push({ top, behavior: options.behavior });
    if (options.behavior === "smooth") {
      animation = { from: scrollTop, to: top, frame: 0 };
      return;
    }
    animation = null;
    scrollTop = top;
  };
  globalThis.requestAnimationFrame = (callback) => {
    frameCallbacks.push(callback);
    return frameCallbacks.length;
  };
  const pageTop = (shell: HTMLElement) =>
    Number(shell.dataset["pageIndex"]) * PAGE_PITCH - scrollTop;
  Object.defineProperty(HTMLElement.prototype, "getBoundingClientRect", {
    configurable: true,
    value(this: HTMLElement) {
      if (this === root) return new DOMRect(0, 0, 1000, VIEWPORT_HEIGHT);
      const shell = this.closest<HTMLElement>(`.${PAGE_CLASS_NAMES.page}`);
      if (!shell) return new DOMRect(0, 0, 0, 0);
      if (shell === this) {
        if (measuring) pageMeasurements++;
        return new DOMRect(0, pageTop(shell), 816, PAGE_HEIGHT);
      }
      return new DOMRect(
        72,
        pageTop(shell) + CONTENT_OFFSET,
        80,
        CONTENT_HEIGHT,
      );
    },
  });

  const { pages, blockLookup } = buildPages();
  renderPages(pages, container, { blockLookup });
  const shells = [
    ...container.querySelectorAll<HTMLElement>(`.${PAGE_CLASS_NAMES.page}`),
  ];
  const paintedIndices = () =>
    shells.flatMap((shell, index) =>
      shell.querySelector(`.${PAGE_CLASS_NAMES.content}`) ? [index] : [],
    );

  const deliverIntersections = (): boolean => {
    let delivered = false;
    for (const observer of [...FakeIntersectionObserver.instances]) {
      const entries: IntersectionObserverEntry[] = [];
      for (const [target, was] of observer.targets) {
        if (!(target instanceof HTMLElement)) continue;
        const top = pageTop(target);
        const isIntersecting =
          top < VIEWPORT_HEIGHT + observer.margin &&
          top + PAGE_HEIGHT > -observer.margin;
        if (was === isIntersecting) continue;
        observer.targets.set(target, isIntersecting);
        entries.push({
          target,
          isIntersecting,
          intersectionRatio: isIntersecting ? 1 : 0,
          boundingClientRect: new DOMRect(0, top, 816, PAGE_HEIGHT),
          intersectionRect: new DOMRect(),
          rootBounds: null,
          time: 0,
        });
      }
      if (entries.length === 0) continue;
      delivered = true;
      measuring = true;
      try {
        observer.callback(entries);
      } finally {
        measuring = false;
      }
    }
    return delivered;
  };

  const blankUnderCentre = () => {
    const centre = scrollTop + VIEWPORT_HEIGHT / 2;
    const shell = shells[Math.floor(centre / PAGE_PITCH)];
    return !shell?.querySelector(`.${PAGE_CLASS_NAMES.content}`);
  };

  /** Run frames until scrolling, animation callbacks and observers are idle. */
  const settle = (): number => {
    let blankFrames = 0;
    for (let frame = 0; frame < MAX_FRAMES; frame++) {
      const moving = animation !== null;
      if (animation) {
        animation.frame++;
        const progress = Math.min(1, animation.frame / SMOOTH_SCROLL_FRAMES);
        scrollTop = animation.from + (animation.to - animation.from) * progress;
        if (progress === 1) animation = null;
      }
      const callbacks = frameCallbacks.splice(0);
      for (const callback of callbacks) callback(frame);
      const delivered = deliverIntersections();
      if (blankUnderCentre()) blankFrames++;
      if (!moving && callbacks.length === 0 && !delivered) return blankFrames;
    }
    throw new Error("The viewport never settled");
  };

  settle();

  /** Navigate with `scrollPagesToPmPosition` and trace what it costs. */
  const jump = (pmPos: number): JumpTrace => {
    const paintedBefore = new Set(paintedIndices());
    const painted = new Set<number>();
    const observer = new MutationObserver((records) => {
      for (const record of records) {
        if (record.addedNodes.length === 0) continue;
        if (!(record.target instanceof HTMLElement)) continue;
        painted.add(Number(record.target.dataset["pageIndex"]));
      }
    });
    for (const shell of shells) observer.observe(shell, { childList: true });
    pageMeasurements = 0;
    scrollCalls.length = 0;
    scrollPagesToPmPosition(container, pmPos);
    const paintedBeforeScroll = paintedIndices().filter(
      (index) => !paintedBefore.has(index),
    );
    const blankFrames = settle();
    for (const record of observer.takeRecords()) {
      if (record.addedNodes.length > 0 && record.target instanceof HTMLElement)
        painted.add(Number(record.target.dataset["pageIndex"]));
    }
    observer.disconnect();
    return {
      paintedBeforeScroll,
      paintedPages: [...painted]
        .filter((index) => !paintedBefore.has(index))
        .sort((a, b) => a - b),
      blankFrames,
      pageMeasurements,
      scrollCalls: [...scrollCalls],
    };
  };

  return { container, shells, paintedIndices, jump };
};

/** Where `scrollEditorElementIntoView` centres content painted on `index`. */
const centredTop = (index: number) =>
  index * PAGE_PITCH + CONTENT_OFFSET - (VIEWPORT_HEIGHT - CONTENT_HEIGHT) / 2;

const setGlobal = (name: string, value: unknown) =>
  Object.defineProperty(globalThis, name, {
    configurable: true,
    writable: true,
    value,
  });

beforeEach(() => {
  FakeIntersectionObserver.instances = [];
  setGlobal("IntersectionObserver", FakeIntersectionObserver);
  Object.defineProperty(window, "innerHeight", {
    configurable: true,
    value: VIEWPORT_HEIGHT,
  });
});

afterEach(() => {
  setGlobal("IntersectionObserver", originalObserver);
  setGlobal("requestAnimationFrame", originalFrame);
  if (originalRect) {
    Object.defineProperty(
      HTMLElement.prototype,
      "getBoundingClientRect",
      originalRect,
    );
  } else {
    Reflect.deleteProperty(HTMLElement.prototype, "getBoundingClientRect");
  }
  if (originalInnerHeight) {
    Object.defineProperty(window, "innerHeight", originalInnerHeight);
  }
  document.body.replaceChildren();
});

describe("scrollPagesToPmPosition over virtualized pages", () => {
  test("a far jump paints the target page and its neighbours first, then lands in one instant move", () => {
    const { jump, paintedIndices } = mountPages();
    expect(paintedIndices()).toEqual([0, 1, 2]);

    const trace = jump(pageStart(60) + 3);
    console.info("FOLIO_OUTLINE_JUMP_MEASUREMENT", JSON.stringify(trace));

    expect(trace.paintedBeforeScroll).toEqual([59, 60, 61]);
    expect(trace.scrollCalls).toEqual([
      { top: centredTop(60), behavior: "instant" },
    ]);
    expect(trace.blankFrames).toBe(0);
    // Only the band virtualization keeps around the target, none in between.
    expect(trace.paintedPages).toEqual([57, 58, 59, 60, 61, 62]);
    expect(paintedIndices()).toEqual([57, 58, 59, 60, 61, 62]);
    expect(trace.pageMeasurements).toBe(0);
  });

  test("a near jump stays smooth unless the reader prefers reduced motion", () => {
    const { jump } = mountPages();
    expect(jump(pageStart(1) + 3).scrollCalls).toEqual([
      { top: centredTop(1), behavior: "smooth" },
    ]);

    const originalMatchMedia = window.matchMedia;
    window.matchMedia = (query: string) => {
      const list = originalMatchMedia.call(window, query);
      Object.defineProperty(list, "matches", {
        value: query === "(prefers-reduced-motion: reduce)",
      });
      return list;
    };
    try {
      expect(jump(pageStart(2) + 3).scrollCalls).toEqual([
        { top: centredTop(2), behavior: "auto" },
      ]);
    } finally {
      window.matchMedia = originalMatchMedia;
    }
  });
});
