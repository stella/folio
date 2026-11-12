/**
 * Where every Playwright config writes its results. CI uploads this one
 * directory when a browser job fails, so a config that writes elsewhere loses
 * its failure evidence; `scripts/playwright-output-dirs.test.ts` holds every
 * config to it.
 */

import path from "node:path";

export const PLAYWRIGHT_RESULTS_ROOT = path.resolve(
  import.meta.dirname,
  "../../test-results",
);

/** A config's own results directory inside the uploaded root. */
export const playwrightOutputDir = (name: string): string =>
  path.join(PLAYWRIGHT_RESULTS_ROOT, name);
