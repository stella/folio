import { expect, test as base, type Page } from "@playwright/test";
import { isCanonicalInputTimer } from "../parity/canonicalTimerOwner";

declare global {
  var __folioCanonicalTimerStacks: Map<number, string> | undefined;
}

/** Runs as an init script before any document or application script. */
const installCanonicalTimerProbe = () => {
  if (globalThis.__folioCanonicalTimerStacks) return;
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
};

/** Every navigation, including reloads, captures application startup timers. */
export const test = base.extend({
  page: async ({ page }, use) => {
    await page.addInitScript(installCanonicalTimerProbe);
    await use(page);
  },
});

/** Read an already instrumented page; observation must never install the probe. */
export const assertCanonicalInputTimersSettled = async (page: Page) => {
  const pending = await page.evaluate(() => {
    if (!globalThis.__folioCanonicalTimerStacks)
      throw new TypeError("Canonical timer instrumentation must be installed before navigation");
    return [...globalThis.__folioCanonicalTimerStacks.values()];
  });
  expect(
    pending.filter(isCanonicalInputTimer),
    "case must start with no canonical input timer",
  ).toEqual([]);
};
