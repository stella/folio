import { expect } from "@playwright/test";
import { test } from "./canonicalTimerProbe";
import fc from "fast-check";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import { checkCanonicalBrowserHistory } from "./canonicalBrowserHistoryOracle";
import {
  CANONICAL_BROWSER_HISTORY_REPLAYS,
  CANONICAL_BROWSER_SAVE_REPLAYS,
  canonicalBrowserTraceArbitrary,
} from "./canonicalBrowserTrace";

for (const { seed, path, kinds } of [
  ...CANONICAL_BROWSER_HISTORY_REPLAYS,
  ...CANONICAL_BROWSER_SAVE_REPLAYS,
]) {
  test(`canonical history replay ${seed} ${path}`, async ({ page }) => {
    const traces = fc.sample(canonicalBrowserTraceArbitrary, { seed, path, numRuns: 1 });
    expect(traces).toHaveLength(1);
    const actions = traces.at(0);
    if (actions === undefined) throw new TypeError("Missing canonical regression trace");
    expect(actions.length).toBeGreaterThan(0);
    expect(actions.map(({ kind }) => kind)).toEqual(kinds);
    await page.goto("/?session=canonical");
    await page.waitForSelector(".layout-page");
    await page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    });
    const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
    // Repeat identical package input to exercise adoption of a fresh owner.
    for (let load = 0; load < 2; load++) {
      const applied = await checkCanonicalBrowserHistory({
        page,
        source: [...new Uint8Array(source)],
        actions,
        missing: createMissingOpBurndown(),
      });
      expect(applied).toBeGreaterThan(0);
    }
  });
}
