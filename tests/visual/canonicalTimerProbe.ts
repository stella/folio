import { expect, test as base, type Page } from "@playwright/test";
import { CANONICAL_INPUT_TIMER_OWNER } from "../../packages/core/src/controller/canonicalInputTimer";

declare global {
  var __folioCanonicalInputTimers: Set<number> | undefined;
}

/** Runs as an init script before any document or application script. */
const installCanonicalTimerProbe = (ownerKey: string) => {
  if (globalThis.__folioCanonicalInputTimers) return;
  const timers = new Set<number>();
  const ownerTag = Symbol.for(ownerKey);
  globalThis.__folioCanonicalInputTimers = timers;
  const schedule = window.setTimeout.bind(window);
  const cancel = window.clearTimeout.bind(window);
  window.setTimeout = (handler, delay, ...args) => {
    // String handlers keep their native evaluation semantics.
    if (typeof handler !== "function") return schedule(handler, delay, ...args);

    const id = schedule(() => {
      timers.delete(id);
      handler.apply(window, args);
    }, delay);
    if (Reflect.get(handler, ownerTag) === true) timers.add(id);
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
    await page.addInitScript(installCanonicalTimerProbe, CANONICAL_INPUT_TIMER_OWNER);
    await use(page);
  },
});

/** Read an already instrumented page; observation must never install the probe. */
export const assertCanonicalInputTimersSettled = async (page: Page) => {
  const pending = await page.evaluate(() => {
    if (!globalThis.__folioCanonicalInputTimers)
      throw new TypeError("Canonical timer instrumentation must be installed before navigation");
    return [...globalThis.__folioCanonicalInputTimers.values()];
  });
  expect(pending, "case must start with no canonical input timer").toEqual([]);
};
