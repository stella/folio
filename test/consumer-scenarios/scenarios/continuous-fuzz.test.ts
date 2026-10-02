/**
 * Continuous fuzzing (support/fuzz-loop.ts), off unless
 * FOLIO_SCENARIO_BUDGET_SECONDS is set: fresh seeded flows from
 * FOLIO_SCENARIO_SEED on, and mutants of the corpus in
 * FOLIO_SCENARIO_CORPUS_DIR plus the checked-in flows, until the budget runs
 * out. Each new failure is shrunk (FOLIO_SCENARIO_SHRINK_SECONDS per
 * failure, 0 to skip) and written to FOLIO_SCENARIO_FAILURES_DIR; the test
 * fails when any flow failed. FOLIO_SCENARIO_FUZZ_STEPS sets a fresh flow's
 * length and FOLIO_SCENARIO_MUTATE_SHARE the share of mutants. With
 * FOLIO_SCENARIO_MINIMIZED_DIR, up to three flows that reach a new structure
 * or step outcome are shrunk into it, as candidates for ../flows. A failure
 * whose fingerprint FOLIO_SCENARIO_KNOWN_FINGERPRINTS lists (the runner sets
 * it from test/known-failure-fingerprints.json) is recorded but does not
 * fail the test.
 */

import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "../support/corpus.ts";
import { reportFuzzHealth } from "../support/fuzz-health.ts";
import { fuzzFor } from "../support/fuzz-loop.ts";

const number = (name: string, fallback: number): number => {
  const parsed = Number(process.env[name]);
  return process.env[name] !== undefined && Number.isFinite(parsed) && parsed >= 0
    ? parsed
    : fallback;
};
const directory = (name: string): string | null => {
  const value = process.env[name];
  return value === undefined || value === "" ? null : value;
};

const BUDGET = number("FOLIO_SCENARIO_BUDGET_SECONDS", 0);
const SEED = Math.floor(number("FOLIO_SCENARIO_SEED", 20_260_926));
const STEPS = Math.max(1, Math.floor(number("FOLIO_SCENARIO_FUZZ_STEPS", 10)));
const SHRINK_SECONDS = number("FOLIO_SCENARIO_SHRINK_SECONDS", 300);
const MINIMIZED_DIR = directory("FOLIO_SCENARIO_MINIMIZED_DIR");
const KNOWN = new Set(
  (process.env["FOLIO_SCENARIO_KNOWN_FINGERPRINTS"] ?? "").split(",").filter(Boolean),
);

if (BUDGET === 0) {
  test("continuous fuzz", { skip: "FOLIO_SCENARIO_BUDGET_SECONDS is not set" }, () => {});
} else {
  test(`continuous fuzz for ${BUDGET}s from seed ${SEED}`, async () => {
    reportFuzzHealth({ status: "started", completed: 0 });
    const result = await fuzzFor({
      seed: SEED,
      seconds: BUDGET,
      steps: STEPS,
      maxSteps: STEPS * 2,
      mutateShare: Math.min(1, number("FOLIO_SCENARIO_MUTATE_SHARE", 0.5)),
      corpusDir: directory("FOLIO_SCENARIO_CORPUS_DIR"),
      corpusSize: Math.floor(number("FOLIO_SCENARIO_CORPUS_SIZE", 400)),
      seeds: readFlowDir(fileURLToPath(new URL("../flows/", import.meta.url))).map(
        ({ flow }) => flow,
      ),
      failuresDir: directory("FOLIO_SCENARIO_FAILURES_DIR"),
      shrink: SHRINK_SECONDS === 0 ? null : { maxAttempts: 200, seconds: SHRINK_SECONDS },
      minimize: MINIMIZED_DIR === null ? null : { dir: MINIMIZED_DIR, max: 3, seconds: 60 },
    });
    console.log(
      `continuous fuzz: ${result.flows} fresh flows, ${result.mutants} mutants, ${result.admitted} new to the corpus (now ${result.corpus}), ${result.minimized} minimized, ${result.failures.length} failing fingerprints`,
    );
    const summary = ({ record, count }: (typeof result.failures)[number]): string =>
      `  ${record.marker.fingerprint} ×${count} ${record.marker.test}: ${record.marker.assertion}\n    Replay: ${record.replays[0]}`;
    // The registry may list a shrunk failure's fingerprint or the one before shrinking.
    const isKnown = ({ record }: (typeof result.failures)[number]): boolean =>
      KNOWN.has(record.marker.fingerprint) ||
      (record.marker.primary !== undefined && KNOWN.has(record.marker.primary));
    const known = result.failures.filter(isKnown);
    if (known.length > 0) {
      console.log([`known failures, not failing the run:`, ...known.map(summary)].join("\n"));
    }
    const fresh = result.failures.filter((failure) => !isKnown(failure));
    const completed = result.completedCases;
    if (completed === 0) {
      reportFuzzHealth({ status: "infrastructure", completed, detail: "No fuzz cases completed" });
      throw new Error("Fuzz infrastructure: no cases completed");
    }
    reportFuzzHealth(
      fresh.length > 0
        ? { status: "finding", completed, detail: `${fresh.length} failing fingerprints` }
        : { status: "passed", completed },
    );
    if (fresh.length > 0) {
      throw new Error(
        [`${fresh.length} failing fingerprint(s):`, ...fresh.map(summary)].join("\n"),
      );
    }
  });
}
