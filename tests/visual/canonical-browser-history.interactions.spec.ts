import { expect, test } from "@playwright/test";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import fc from "fast-check";
import { Result } from "better-result";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import { checkCanonicalBrowserHistory } from "./canonicalBrowserHistoryOracle";
import {
  CANONICAL_BROWSER_HISTORY_REPLAYS,
  CANONICAL_BROWSER_SAVE_REPLAYS,
  CANONICAL_BROWSER_REDO_REPLAYS,
  canonicalBrowserTraceArbitrary,
} from "./canonicalBrowserTrace";
import type { CanonicalFuzzObservation } from "../parity/canonicalFuzzErrors";
import {
  failureMarker,
  logFailureMarker,
  writeFailureRecord,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import { canonicalOracleFailureRecord } from "../parity/canonicalOracleFailure";

const replaySet = process.env["FOLIO_CANONICAL_REPLAY_SET"];
if (replaySet !== undefined && replaySet !== "redo")
  throw new TypeError("Unknown canonical replay set");
const replays =
  replaySet === "redo"
    ? CANONICAL_BROWSER_REDO_REPLAYS
    : [...CANONICAL_BROWSER_HISTORY_REPLAYS, ...CANONICAL_BROWSER_SAVE_REPLAYS];

const repetitions = replaySet === "redo" ? 25 : 1;
for (const { seed, path, kinds } of replays) {
  for (let repetition = 0; repetition < repetitions; repetition++) {
    test(`canonical history replay ${seed} ${path} repetition ${repetition + 1}/${repetitions}`, async ({
      page,
    }, info) => {
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
      const observations: CanonicalFuzzObservation[] = [];
      const outcome = await Result.tryPromise({
        try: async () => {
          // Repeat identical package input to exercise adoption of a fresh owner.
          for (let load = 0; load < 2; load++) {
            const applied = await checkCanonicalBrowserHistory({
              page,
              source: [...new Uint8Array(source)],
              actions,
              missing: createMissingOpBurndown(),
              observations,
            });
            expect(applied).toBeGreaterThan(0);
          }
        },
        catch: (cause: unknown) => cause,
      });
      const observationPath = info.outputPath("canonical-history-observations.json");
      await mkdir(dirname(observationPath), { recursive: true });
      await writeFile(
        observationPath,
        JSON.stringify({
          seed,
          path,
          repetition,
          status: outcome.isErr() ? "failed" : "passed",
          actions,
          observations,
        }),
      );
      await info.attach("canonical-history-observations", {
        path: observationPath,
        contentType: "application/json",
      });
      if (outcome.isErr()) {
        const failure = outcome.error;
        const marker = failureMarker({
          test: "canonical browser history replay",
          seed,
          path,
          repro:
            "bunx playwright test --project=interactions tests/visual/canonical-browser-history.interactions.spec.ts --workers=1",
          failure,
          flow: actions.map(({ kind }) => kind).join(" → "),
        });
        logFailureMarker(marker);
        const artifact = writeFailureRecord(
          process.env["FOLIO_FUZZ_FAILURES_DIR"] ?? "fuzz-artifacts/canonical/findings",
          canonicalOracleFailureRecord({ marker, failure, flow: actions }),
        );
        await info.attach("canonical-history-failure", {
          path: artifact,
          contentType: "application/json",
        });
        throw failure;
      }
    });
  }
}
