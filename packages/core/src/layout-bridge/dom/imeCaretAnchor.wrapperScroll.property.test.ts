/**
 * The hidden editor's host sits in a clipping wrapper that the browser scrolls,
 * by any amount and at any time, to reveal the native caret the anchor's
 * translation moved outside it; the scroll event follows later. Over any
 * sequence of such scrolls, their events and syncs, the native caret must stay
 * on the painted caret the last sync anchored it to.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import {
  HIDDEN_EDITOR_WRAPPER_CLASS,
  syncImeCaretAnchor,
} from "./imeCaretAnchor";

setDefaultTimeout(propertyTestTimeout(30_000));

type Point = { left: number; top: number };

/** Offsets the way a browser applies the host's inline transform. */
const appliedOffset = (transform: string): { dx: number; dy: number } => {
  if (transform === "") return { dx: 0, dy: 0 };
  const match = /^translate3d\((-?[\d.]+)px, (-?[\d.]+)px, 0(?:px)?\)$/u.exec(
    transform,
  );
  if (!match) throw new Error(`unexpected transform ${transform}`);
  return { dx: Number(match[1]), dy: Number(match[2]) };
};

/** A wrapper whose scroll events wait until the browser delivers them. */
const scrollingWrapper = () => {
  const listeners: (() => void)[] = [];
  const wrapper = {
    scrollLeft: 0,
    scrollTop: 0,
    addEventListener: (_type: "scroll", listener: () => void) => {
      listeners.push(listener);
    },
    closest: (selectors: string) =>
      selectors === `.${HIDDEN_EDITOR_WRAPPER_CLASS}` ? wrapper : null,
  };
  return {
    wrapper,
    deliverScroll: () => {
      for (const listener of listeners) listener();
    },
    listeners: () => listeners.length,
  };
};

// Layout positions are multiples of 1/64 px, so sums stay exact.
const coordinate = fc
  .integer({ min: -20_000 * 64, max: 20_000 * 64 })
  .map((units) => units / 64);
const point = fc.record({ left: coordinate, top: coordinate });
const scrollOffset = fc.integer({ min: 0, max: 40_000 });

const steps = fc.array(
  fc.oneof(
    fc.record({
      kind: fc.constant("sync" as const),
      caret: point,
      visible: point,
    }),
    fc.record({
      kind: fc.constant("scroll" as const),
      left: scrollOffset,
      top: scrollOffset,
    }),
    fc.constant({ kind: "event" as const }),
  ),
  { minLength: 1, maxLength: 16 },
);

describe("syncImeCaretAnchor in a scrolling wrapper", () => {
  test("keeps the native caret on the painted caret whatever the wrapper is scrolled by", () => {
    fc.assert(
      fc.property(
        fc.record({ left: scrollOffset, top: scrollOffset }),
        steps,
        (initialScroll, run) => {
          const { wrapper, deliverScroll } = scrollingWrapper();
          wrapper.scrollLeft = initialScroll.left;
          wrapper.scrollTop = initialScroll.top;
          const host = { style: { transform: "" }, parentElement: wrapper };
          // Where the caret is in the document, before any translation or
          // scroll moves it on screen.
          let caret: Point = { left: 0, top: 0 };
          const onScreen = (): Point => {
            const { dx, dy } = appliedOffset(host.style.transform);
            return {
              left: caret.left + dx - wrapper.scrollLeft,
              top: caret.top + dy - wrapper.scrollTop,
            };
          };
          const editorView = {
            composing: false,
            hasFocus: () => true,
            state: { selection: { empty: true, head: 1 } },
            coordsAtPos: onScreen,
          };
          let anchoredTo: Point | null = null;
          for (const step of run) {
            if (step.kind === "sync") {
              caret = step.caret;
              anchoredTo = syncImeCaretAnchor({
                hiddenHost: host,
                editorView,
                visibleCaret: step.visible,
              })
                ? step.visible
                : null;
            } else if (step.kind === "scroll") {
              wrapper.scrollLeft = step.left;
              wrapper.scrollTop = step.top;
            } else {
              deliverScroll();
              // Once the browser has told the page about its scroll, the
              // caret is back where the last sync put it.
              if (anchoredTo !== null) {
                const native = onScreen();
                expect(Math.abs(native.left - anchoredTo.left)).toBeLessThan(
                  0.5 + 1e-9,
                );
                expect(Math.abs(native.top - anchoredTo.top)).toBeLessThan(
                  0.5 + 1e-9,
                );
              }
            }
            if (step.kind === "sync" && anchoredTo !== null) {
              const native = onScreen();
              expect(Math.abs(native.left - anchoredTo.left)).toBeLessThan(
                0.5 + 1e-9,
              );
              expect(Math.abs(native.top - anchoredTo.top)).toBeLessThan(
                0.5 + 1e-9,
              );
            }
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("listens to a wrapper once, however often it syncs", () => {
    const { wrapper, listeners } = scrollingWrapper();
    const host = { style: { transform: "" }, parentElement: wrapper };
    const editorView = {
      composing: false,
      hasFocus: () => true,
      state: { selection: { empty: true, head: 1 } },
      coordsAtPos: () => ({ left: 0, top: 0 }),
    };
    for (let sync = 0; sync < 5; sync += 1)
      syncImeCaretAnchor({
        hiddenHost: host,
        editorView,
        visibleCaret: { left: sync, top: sync },
      });
    expect(listeners()).toBe(1);
  });

  test("leaves a host outside a wrapper as it was", () => {
    const host = {
      style: { transform: "" },
      parentElement: { closest: () => null },
    };
    expect(
      syncImeCaretAnchor({
        hiddenHost: host,
        editorView: {
          composing: false,
          hasFocus: () => true,
          state: { selection: { empty: true, head: 1 } },
          coordsAtPos: () => ({ left: 10, top: 20 }),
        },
        visibleCaret: { left: 30, top: 50 },
      }),
    ).toBe(true);
    expect(host.style.transform).toBe("translate3d(20px, 30px, 0)");
  });
});
