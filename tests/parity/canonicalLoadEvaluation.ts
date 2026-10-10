import type { Page } from "@playwright/test";
import type {} from "./canonicalBridge";

/** CDP awaiting a promise does not keep it alive across Chromium collection. */
export const evaluateCanonicalLoad = async (page: Page, bytes: number[]) => {
  // A non-thenable box gives Playwright a strong handle to the exact operation.
  const operation = await page.evaluateHandle(
    (source) => ({
      pending: globalThis.__folioCanonical?.load(source),
    }),
    bytes,
  );
  try {
    return await operation.evaluate(({ pending }) => pending);
  } finally {
    await operation.dispose();
  }
};
