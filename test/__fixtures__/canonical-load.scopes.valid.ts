import type { Page } from "@playwright/test";
import type {} from "../../tests/parity/canonicalBridge";

export const loadUnrelatedParameter = (page: Page) =>
  page.evaluate(() => {
    const bridge = globalThis.__folioCanonical;
    const invokeUnrelated = (bridge: { load: () => boolean }) => bridge.load();
    return { canonical: bridge !== undefined, unrelated: invokeUnrelated({ load: () => true }) };
  });

export const loadUnrelatedDestructuredParameter = (page: Page) =>
  page.evaluate(() => {
    const load = globalThis.__folioCanonical?.load;
    const invokeUnrelated = ({ load }: { load: () => boolean }) => load();
    return { canonical: load !== undefined, unrelated: invokeUnrelated({ load: () => true }) };
  });

export const loadShadowedCanonicalName = (page: Page) =>
  page.evaluate(() => {
    const invokeUnrelated = (__folioCanonical: { load: () => boolean }) => __folioCanonical.load();
    return invokeUnrelated({ load: () => true });
  });

export const loadUnrelatedLocal = (page: Page) =>
  page.evaluate(() => {
    const __folioCanonical = { load: () => true };
    return __folioCanonical.load();
  });
