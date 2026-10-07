import { expect, type Page } from "@playwright/test";
import { isCanonicalInputTimer } from "../parity/canonicalTimerOwner";

declare global {
  var __folioCanonicalTimerStacks: Map<number, string> | undefined;
}

/** Capture the scheduling stack; classify ownership outside the browser. */
export const assertCanonicalInputTimersSettled = async (page: Page) => {
  const pending = await page.evaluate(() => {
    if (!globalThis.__folioCanonicalTimerStacks) {
      const timers = new Map<number, string>();
      globalThis.__folioCanonicalTimerStacks = timers;
      const schedule = window.setTimeout.bind(window);
      const cancel = window.clearTimeout.bind(window);
      window.setTimeout = (handler, delay, ...args) => {
        // String handlers keep their native evaluation semantics.
        if (typeof handler !== "function") return schedule(handler, delay, ...args);
        const stack = new Error().stack;
        if (!stack) throw new TypeError("Timer capture has no scheduling stack");
        const id = schedule(() => {
          timers.delete(id);
          handler.apply(window, args);
        }, delay);
        timers.set(id, stack);
        return id;
      };
      window.clearTimeout = (id) => {
        if (id !== undefined) timers.delete(id);
        cancel(id);
      };
    }
    return [...globalThis.__folioCanonicalTimerStacks.values()];
  });
  expect(
    pending.filter(isCanonicalInputTimer),
    "case must start with no canonical input timer",
  ).toEqual([]);
};
