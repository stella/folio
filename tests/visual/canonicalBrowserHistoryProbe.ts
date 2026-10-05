import type { Page } from "@playwright/test";
import type {} from "../parity/canonicalBridge";
import type {} from "../parity/canonicalHistoryObservation";
import type { CanonicalFuzzPhase } from "../parity/canonicalFuzzErrors";

/** Observe native delivery without changing editor handlers or focus. */
export const installCanonicalHistoryProbe = (page: Page) =>
  page.evaluate(() => {
    if (globalThis.__folioCanonicalHistoryKeys !== undefined) return;
    globalThis.__folioCanonicalHistoryKeys = [];
    const record = (event: KeyboardEvent, propagation: "capture" | "bubble") => {
      const bridge = globalThis.__folioCanonical;
      if (!bridge) return;
      globalThis.__folioCanonicalHistoryKeys?.push({
        propagation,
        key: event.key,
        control: event.ctrlKey,
        meta: event.metaKey,
        shift: event.shiftKey,
        defaultPrevented: event.defaultPrevented,
        target: bridge.ownsKeyTarget(event.target) ? "editor" : "other",
        state: bridge.historyState("input"),
      });
    };
    document.addEventListener("keydown", (event) => record(event, "capture"), true);
    document.addEventListener("keydown", (event) => record(event, "bubble"));
  });

export const beginCanonicalHistoryObservation = (page: Page, phase: CanonicalFuzzPhase) =>
  page.evaluate((current) => {
    globalThis.__folioCanonicalFuzzPhase = current;
    globalThis.__folioCanonicalHistoryKeys?.splice(0);
    const bridge = globalThis.__folioCanonical;
    if (!bridge) throw new TypeError("Canonical history bridge unavailable");
    return {
      before: bridge.historyState("version"),
      keys: [],
      capture: { status: "complete" as const },
    };
  }, phase);

export const collectCanonicalHistoryObservation = (page: Page) =>
  page.evaluate(() => {
    const bridge = globalThis.__folioCanonical;
    const keys = globalThis.__folioCanonicalHistoryKeys;
    const errors = globalThis.__folioCanonicalFuzzErrors;
    if (!errors) throw new TypeError("Canonical error sink unavailable");
    // This is a browser boundary: report observer failure while retaining the independent error sink.
    const capture = (() => {
      try {
        if (!bridge || !keys) throw new TypeError("Canonical history probe unavailable");
        return {
          status: "complete",
          after: bridge.historyState("version"),
          keys: keys.splice(0),
        } as const;
      } catch (cause) {
        return {
          status: "unavailable",
          message: cause instanceof Error ? cause.message : String(cause),
        } as const;
      }
    })();
    return { capture, errors: errors.splice(0) };
  });
