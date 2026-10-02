import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, describe, expect, test } from "bun:test";
import {
  getEditorScrollRoot,
  scrollEditorBy,
  scrollEditorElementIntoView,
  scrollEditorTo,
} from "./editorScrollRoot";

afterAll(() => GlobalRegistrator.unregister());

const fixture = () => {
  const host = document.createElement("div");
  const root = document.createElement("div");
  root.setAttribute("data-folio-scroll", "");
  const viewport = document.createElement("div");
  const target = document.createElement("div");
  host.append(root);
  root.append(viewport);
  viewport.append(target);
  Object.defineProperties(root, {
    clientTop: { value: 2, configurable: true },
    clientHeight: { value: 300, configurable: true },
    clientWidth: { value: 800, configurable: true },
  });
  root.getBoundingClientRect = () => new DOMRect(0, 100, 800, 304);
  target.getBoundingClientRect = () => new DOMRect(0, 702, 400, 40);
  host.scrollTop = 77;
  root.scrollTop = 200;
  return { host, root, viewport, target };
};

describe("editor scroll root ownership", () => {
  test("resolves the same marked root from every descendant even without overflow", () => {
    const { root, viewport, target } = fixture();
    for (const element of [root, viewport, target]) {
      expect(getEditorScrollRoot(element)).toBe(root);
    }
    expect(getEditorScrollRoot(null)).toBeNull();
    expect(getEditorScrollRoot(document.createElement("div"))).toBeNull();
  });

  test("absolute and relative scroll writes move only the editor root", () => {
    const { root, host, target, viewport } = fixture();
    scrollEditorTo(target, { top: 450, left: 25, behavior: "instant" });
    scrollEditorBy(viewport, 12);
    expect(root.scrollTop).toBe(462);
    expect(root.scrollLeft).toBe(25);
    expect(viewport.scrollTop).toBe(0);
    expect(host.scrollTop).toBe(77);
  });

  test("alignment stays in root pixels across target scale and scroll offsets", () => {
    for (const scrollTop of [0, 200, 1000]) {
      for (const scale of [0.5, 1, 2]) {
        for (const block of ["start", "center", "end", "nearest"] as const) {
          const { root, host, viewport, target } = fixture();
          root.scrollTop = scrollTop;
          target.getBoundingClientRect = () => new DOMRect(0, 702, 400, 40 * scale);
          scrollEditorElementIntoView(target, { block, margin: 24, behavior: "instant" });
          const expected = {
            start: scrollTop + 600 - 24,
            center: scrollTop + 600 - (300 - 40 * scale) / 2,
            end: scrollTop + 600 + 40 * scale - 300 + 24,
            nearest: scrollTop + 600 + 40 * scale - 300 + 24,
          };
          expect(root.scrollTop).toBe(expected[block]);
          expect(host.scrollTop).toBe(77);
          expect(viewport.scrollTop).toBe(0);
        }
      }
    }
  });

  test("nearest leaves an already visible target and all ancestors still", () => {
    const { root, host, target } = fixture();
    target.getBoundingClientRect = () => new DOMRect(0, 160, 400, 40);
    scrollEditorElementIntoView(target, { block: "nearest", margin: 40, behavior: "instant" });
    expect(root.scrollTop).toBe(200);
    expect(host.scrollTop).toBe(77);
  });

  test("horizontal nearest moves only the root and preserves oversized intersections", () => {
    for (const [left, width, expected] of [
      [850, 100, 175],
      [-50, 100, 0],
      [-50, 900, 25],
      [850, 900, 875],
    ] as const) {
      const { root, host, viewport, target } = fixture();
      root.scrollLeft = 25;
      target.getBoundingClientRect = () => new DOMRect(left, 160, width, 40);
      scrollEditorElementIntoView(target, { block: "nearest", behavior: "instant" });
      expect(root.scrollLeft).toBe(expected);
      expect(host.scrollTop).toBe(77);
      expect(viewport.scrollLeft).toBe(0);
    }
  });

  test("nearest preserves a target spanning both visible vertical edges", () => {
    const { root, target } = fixture();
    target.getBoundingClientRect = () => new DOMRect(0, 50, 400, 500);
    scrollEditorElementIntoView(target, { block: "nearest", margin: 40, behavior: "instant" });
    expect(root.scrollTop).toBe(200);
  });

  test("unmarked hosts cannot become scroll targets", () => {
    const { host, root, target } = fixture();
    root.removeAttribute("data-folio-scroll");
    scrollEditorTo(target, { top: 500 });
    scrollEditorBy(target, 100);
    scrollEditorElementIntoView(target);
    expect(root.scrollTop).toBe(200);
    expect(host.scrollTop).toBe(77);
  });
});
