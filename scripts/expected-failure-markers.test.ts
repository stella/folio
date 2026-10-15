import { afterEach, beforeEach, describe, expect, spyOn, test } from "bun:test";
import { reportScenarioFailure } from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  FINDING_SYMPTOMS,
  settleExpectedFailure,
} from "../test/consumer-scenarios/support/known-issues";
import { extractFailureMarkers } from "./failure-fingerprints";

// The checked-in flow behind BATCH_INSERT_BEFORE_ROW_PENDING_DELETION, as a
// merge-group run logged it while the flow passed as an expected failure.
const FLOW = {
  test: "consumer flow stories / suggested",
  seed: 774176929,
  repro:
    "bun scripts/consumer-scenarios.ts --only '^checked-in flow batch-after-table-row-delete\\.json' -- flow-corpus.test.ts",
};
const KNOWN =
  "[batchSequential] the batch (suggested: insertBeforeBlock, insertBeforeBlock, insertTableRow) leaves otherwise than its operations one at a time:\n" +
  "  op-1 refused one at a time: pendingDeletion";
const OTHER =
  "[batchSequential] the batch (suggested: insertBeforeBlock, insertBeforeBlock, insertTableRow) leaves otherwise than its operations one at a time:\n" +
  "  op-1 applied one at a time, refused in the batch";

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
    await settleExpectedFailure(
      "BATCH_INSERT_BEFORE_ROW_PENDING_DELETION",
      "checked-in flow batch-after-table-row-delete.json",
      FINDING_SYMPTOMS.BATCH_INSERT_BEFORE_ROW_PENDING_DELETION,
      failWith(KNOWN),
    );
    expect(markers()).toEqual([]);
  });

  test("a failure with another symptom keeps its marker and fails", async () => {
    const settled = settleExpectedFailure(
      "BATCH_INSERT_BEFORE_ROW_PENDING_DELETION",
      "checked-in flow batch-after-table-row-delete.json",
      FINDING_SYMPTOMS.BATCH_INSERT_BEFORE_ROW_PENDING_DELETION,
      failWith(OTHER),
    );
    await expect(settled).rejects.toThrow("refused in the batch");
    expect(markers().map((marker) => ({ test: marker.test, seed: marker.seed }))).toEqual([
      { test: FLOW.test, seed: FLOW.seed },
    ]);
  });

  test("a body that passes fails as no longer reproducing, with no marker", async () => {
    const settled = settleExpectedFailure(
      "BATCH_INSERT_BEFORE_ROW_PENDING_DELETION",
      "checked-in flow batch-after-table-row-delete.json",
      FINDING_SYMPTOMS.BATCH_INSERT_BEFORE_ROW_PENDING_DELETION,
      () => {},
    );
    await expect(settled).rejects.toThrow("no longer reproduces");
    expect(markers()).toEqual([]);
  });

  test("markers logged outside an expected failure are logged at once", () => {
    expect(failWith(KNOWN)).toThrow("pendingDeletion");
    expect(markers()).toHaveLength(1);
  });
});
