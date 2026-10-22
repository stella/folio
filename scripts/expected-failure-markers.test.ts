import { afterEach, beforeEach, describe, expect, spyOn, test } from "bun:test";
import { reportScenarioFailure } from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  FINDING_SYMPTOMS,
  matchesFailureSymptom,
  settleExpectedFailure,
} from "../test/consumer-scenarios/support/known-issues";
import { extractFailureMarkers } from "./failure-fingerprints";

// A representative scenario failure, as a merge-group run logged it while the
// scenario passed as an expected failure.
const FLOW = {
  test: "consumer known issues / suggested",
  seed: 774176929,
  repro:
    "bun scripts/consumer-scenarios.ts --only '^a paragraph inserted after a block with a pending tracked merge'",
};
const KNOWN = "accepting glues the inserted paragraph onto the merged one";
const OTHER =
  "[batchSequential] the batch (suggested: insertBeforeBlock, insertBeforeBlock, insertTableRow) leaves otherwise than its operations one at a time:\n" +
  "  op-1 applied one at a time, refused in the batch";
const FINDING = "INSERT_AFTER_PENDING_MERGE";
const SYMPTOM = FINDING_SYMPTOMS[FINDING];
const NAME = "insert after a pending merge";

const failWith = (message: string) => () => {
  reportScenarioFailure({ ...FLOW, failure: new Error(message) });
};

describe("expected failures and the failure markers fuzz reporters file", () => {
  let logged: string[] = [];
  let restore = () => {};
  beforeEach(() => {
    logged = [];
    const spy = spyOn(console, "error").mockImplementation((...parts: unknown[]) => {
      logged.push(parts.map(String).join(" "));
    });
    restore = () => spy.mockRestore();
  });
  afterEach(() => restore());

  const markers = () => extractFailureMarkers(logged.join("\n"));

  test("a failure with the finding's symptom logs no marker", async () => {
    await settleExpectedFailure(FINDING, NAME, SYMPTOM, failWith(KNOWN));
    expect(markers()).toEqual([]);
  });

  test("a failure with another symptom keeps its marker and fails", async () => {
    const settled = settleExpectedFailure(FINDING, NAME, SYMPTOM, failWith(OTHER));
    await expect(settled).rejects.toThrow("refused in the batch");
    expect(markers().map((marker) => ({ test: marker.test, seed: marker.seed }))).toEqual([
      { test: FLOW.test, seed: FLOW.seed },
    ]);
  });

  test("a body that passes fails as no longer reproducing, with no marker", async () => {
    const settled = settleExpectedFailure(FINDING, NAME, SYMPTOM, () => {});
    await expect(settled).rejects.toThrow("no longer reproduces");
    expect(markers()).toEqual([]);
  });

  test("an unknown finding fails with its name", () => {
    expect(() => matchesFailureSymptom("REMOVED_FINDING", undefined, new Error(KNOWN))).toThrow(
      "Unknown expected-failure finding: REMOVED_FINDING",
    );
  });

  test("markers logged outside an expected failure are logged at once", () => {
    expect(failWith(KNOWN)).toThrow("accepting glues");
    expect(markers()).toHaveLength(1);
  });
});
