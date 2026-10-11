import type { Page } from "@playwright/test";
import type {} from "../../tests/parity/canonicalBridge";
import { evaluateCanonicalLoad } from "../../tests/parity/canonicalLoadEvaluation";

export const loadWithOwner = (page: Page, bytes: number[]) => evaluateCanonicalLoad(page, bytes);

export const decorateLoad = (page: Page) =>
  page.evaluate(() => {
    const bridge = globalThis.__folioCanonical;
    if (!bridge) return;
    const load = bridge.load;
    bridge.load = (bytes) => load(bytes);
  });

export const evaluateOtherBridgeMethod = (page: Page) =>
  page.evaluate(() => globalThis.__folioCanonical?.snapshot());
