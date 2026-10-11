import type { Page } from "@playwright/test";
import type {} from "../../tests/parity/canonicalBridge";

export const loadSiblingParameter = (page: Page) =>
  page.evaluate(() => {
    const canonical = (() => {
      const bridge = globalThis.__folioCanonical;
      return bridge !== undefined;
    })();
    const invokeUnrelated = (bridge: { load: () => boolean }) => bridge.load();
    return { canonical, unrelated: invokeUnrelated({ load: () => true }) };
  });

export const loadSiblingDestructuredParameter = (page: Page) =>
  page.evaluate(() => {
    const canonical = (() => {
      const load = globalThis.__folioCanonical?.load;
      return load !== undefined;
    })();
    const invokeUnrelated = ({ load }: { load: () => boolean }) => load();
    return { canonical, unrelated: invokeUnrelated({ load: () => true }) };
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
