import { expect, type Page } from "@playwright/test";
import fc from "fast-check";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import { evaluateCanonicalPage } from "./canonicalPageNavigation";
import {
  checkCanonicalBrowserHistory,
  initializeCanonicalBrowserHistory,
} from "./canonicalBrowserHistoryOracle";
import {
  CANONICAL_BROWSER_HISTORY_REPLAYS,
  CANONICAL_BROWSER_SAVE_REPLAYS,
  canonicalBrowserTraceArbitrary,
} from "./canonicalBrowserTrace";

type CanonicalHistoryReplay =
  | (typeof CANONICAL_BROWSER_HISTORY_REPLAYS)[number]
  | (typeof CANONICAL_BROWSER_SAVE_REPLAYS)[number];

export const canonicalHistoryReplayActions = ({ seed, path, kinds }: CanonicalHistoryReplay) => {
  const traces = fc.sample(canonicalBrowserTraceArbitrary, { seed, path, numRuns: 1 });
  expect(traces).toHaveLength(1);
  const actions = traces.at(0);
  if (actions === undefined) throw new TypeError("Missing canonical regression trace");
  expect(actions.length).toBeGreaterThan(0);
  expect(actions.map(({ kind }) => kind)).toEqual(kinds);
  if (seed === 431 && path === "8")
    expect(actions).toEqual([
      { kind: "imeReplacement", updates: ["shall", "café 東京 é"], completion: "cancel" },
    ]);
  return actions;
};

export const runCanonicalHistoryReplay = async (page: Page, replay: CanonicalHistoryReplay) => {
  const actions = canonicalHistoryReplayActions(replay);
  await page.goto("/?session=canonical");
  await page.waitForSelector(".layout-page");
  await evaluateCanonicalPage(page, () =>
    page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    }),
  );
  const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
  await initializeCanonicalBrowserHistory(page, [...new Uint8Array(source)]);
  // Repeat identical package input to exercise adoption of a fresh owner.
  for (let load = 0; load < 2; load++) {
    const applied = await checkCanonicalBrowserHistory({
      page,
      source: [...new Uint8Array(source)],
      actions,
      missing: createMissingOpBurndown(),
    });
    if (
      actions.every((action) => action.kind === "imeReplacement" && action.completion === "cancel")
    )
      expect(applied).toBe(0);
    else expect(applied).toBeGreaterThan(0);
  }
};
