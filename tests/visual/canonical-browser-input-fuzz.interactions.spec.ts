import { expect, test } from "@playwright/test";
import fc from "fast-check";
import { appendFileSync } from "node:fs";

import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import {
  failureMarker,
  logFailureMarker,
  shellQuote,
  writeFailureRecord,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import { parseBrowserInputTraceConfig } from "./browserInputTrace";
import { canonicalBrowserTraceArbitrary } from "./canonicalBrowserTrace";
import { checkCanonicalBrowserHistory } from "./canonicalBrowserHistoryOracle";
import { canonicalOracleFailureRecord } from "../parity/canonicalOracleFailure";
import type {} from "../parity/canonicalBridge";
import type {} from "../parity/canonicalFuzzErrors";

const config = parseBrowserInputTraceConfig(
  process.env,
  process.env["FOLIO_FUZZ_LANE"] === "nightly" ? "nightly" : "pullRequest",
);
for (const seed of config.seeds) {
  test(`canonical seed ${seed}: projection, exact history and save/reopen`, async ({
    page,
  }, info) => {
    test.setTimeout(600_000);
    const missing = createMissingOpBurndown();
    let applied = 0;
    let completed = 0;
    await page.goto("/?session=canonical");
    await page.waitForSelector(".layout-page");
    await page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    });
    const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
    const verdict = await fc.check(
      fc.asyncProperty(canonicalBrowserTraceArbitrary, async (actions) => {
        applied += await checkCanonicalBrowserHistory({
          page,
          source: [...new Uint8Array(source)],
          actions,
          missing,
        });
        completed++;
      }),
      {
        seed,
        numRuns: config.runs,
        endOnFailure: false,
        ...(process.env["FOLIO_FUZZ_PATH"] === undefined
          ? {}
          : { path: process.env["FOLIO_FUZZ_PATH"] }),
      },
    );
    await info.attach("canonical-missing-ops", {
      body: JSON.stringify({
        seed,
        applied,
        completed,
        missing: missing.rows(),
        verdict: verdict.failed ? "failure" : "completed",
      }),
      contentType: "application/json",
    });
    const summaryPath = process.env["GITHUB_STEP_SUMMARY"];
    if (summaryPath)
      appendFileSync(
        summaryPath,
        `\nCanonical seed ${seed}: ${completed} completed traces, ${applied} applied edits\n\n${missing.markdown()}\n`,
      );
    console.log(
      `Canonical seed ${seed}: ${completed} completed traces, ${applied} applied edits\n${missing.markdown()}`,
    );
    if (verdict.failed) {
      const flow = verdict.counterexample?.at(0);
      const failure = verdict.errorInstance;
      const detail =
        failure instanceof Error ? (failure.stack ?? failure.message) : fc.stringify(failure);
      const message = failure instanceof Error ? failure.message : detail;
      if (
        flow !== undefined &&
        failure !== undefined &&
        failure !== null &&
        message.trim() !== "" &&
        message !== "undefined"
      ) {
        const repro = `FOLIO_FUZZ_LANE=nightly FOLIO_FUZZ_SEEDS=${verdict.seed} FOLIO_FUZZ_RUNS=${config.runs}${verdict.counterexamplePath === null ? "" : ` FOLIO_FUZZ_PATH=${shellQuote(verdict.counterexamplePath)}`} bunx playwright test --project=browser-fuzzer tests/visual/canonical-browser-input-fuzz.interactions.spec.ts --workers=1`;
        const marker = failureMarker({
          test: "canonical browser input: projection, exact history and save/reopen",
          seed: verdict.seed,
          path: verdict.counterexamplePath,
          repro,
          failure,
          flow: flow.map(({ kind }) => kind).join(" → "),
        });
        logFailureMarker(marker);
        const record = canonicalOracleFailureRecord({ marker, failure, flow });
        const artifact = writeFailureRecord(
          process.env["FOLIO_FUZZ_FAILURES_DIR"] ?? "fuzz-artifacts/canonical/findings",
          record,
        );
        await info.attach("canonical-failure-record", {
          path: artifact,
          contentType: "application/json",
        });
      }
      throw new Error(
        `seed=${verdict.seed} path=${verdict.counterexamplePath} trace=${JSON.stringify(flow)}\n${detail}`,
        { cause: failure },
      );
    }
    expect(completed).toBeGreaterThan(0);
  });
}
