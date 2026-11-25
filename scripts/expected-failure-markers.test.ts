import { afterEach, beforeEach, describe, expect, spyOn, test } from "bun:test";

import { reportScenarioFailure } from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  FINDING_SYMPTOMS,
  logRecordedFinding,
  matchesFailureSymptom,
  recordedFindingOf,
  settleExpectedFailure,
} from "../test/consumer-scenarios/support/known-issues";
import { extractFailureMarkers } from "./failure-fingerprints";

// A representative scenario failure, as a merge-group run logged it while the
// scenario passed as an expected failure.
const FLOW = {
  test: "consumer known issues / suggested",
  seed: 774176929,
  repro:
    "bun scripts/consumer-scenarios.ts --only '^docxToMarkdown reads a text box'",
};
const KNOWN = "docxToMarkdown writes no text-box paragraph";
const OTHER =
  "[batchSequential] the batch (suggested: insertBeforeBlock, insertBeforeBlock, insertTableRow) leaves otherwise than its operations one at a time:\n" +
  "  op-1 applied one at a time, refused in the batch";
const FINDING = "MARKDOWN_DROPS_TEXT_BOX";
const SYMPTOM = FINDING_SYMPTOMS[FINDING];
const NAME = "markdown reads a text box";

const failWith = (message: string) => () => {
  reportScenarioFailure({ ...FLOW, failure: new Error(message) });
};

describe("expected failures and the failure markers fuzz reporters file", () => {
  let logged: string[] = [];
  let restore = () => {};
  beforeEach(() => {
    logged = [];
    const spy = spyOn(console, "error").mockImplementation(
      (...parts: unknown[]) => {
        logged.push(parts.map(String).join(" "));
      },
    );
    restore = () => spy.mockRestore();
  });
  afterEach(() => restore());

  const markers = () => extractFailureMarkers(logged.join("\n"));

  test("a failure with the finding's symptom logs no marker", async () => {
    await settleExpectedFailure(FINDING, NAME, SYMPTOM, failWith(KNOWN));
    expect(markers()).toEqual([]);
  });

  test("a failure with another symptom keeps its marker and fails", async () => {
    const settled = settleExpectedFailure(
      FINDING,
      NAME,
      SYMPTOM,
      failWith(OTHER),
    );
    await expect(settled).rejects.toThrow("refused in the batch");
    expect(
      markers().map((marker) => ({ test: marker.test, seed: marker.seed })),
    ).toEqual([{ test: FLOW.test, seed: FLOW.seed }]);
  });

  test("a body that passes fails as no longer reproducing, with no marker", async () => {
    const settled = settleExpectedFailure(FINDING, NAME, SYMPTOM, () => {});
    await expect(settled).rejects.toThrow("no longer reproduces");
    expect(markers()).toEqual([]);
  });

  test("an unknown finding fails with its name", () => {
    expect(() =>
      matchesFailureSymptom("REMOVED_FINDING", undefined, new Error(KNOWN)),
    ).toThrow("Unknown expected-failure finding: REMOVED_FINDING");
  });

  test("markers logged outside an expected failure are logged at once", () => {
    expect(failWith(KNOWN)).toThrow("writes no text-box paragraph");
    expect(markers()).toHaveLength(1);
  });
});

// A generated flow's failure as support/fuzz.ts throws it: the flow's steps,
// then the oracle's message.
const generatedFailure = (message: string) =>
  new Error(
    [
      "collisions flow (targeted) with seed 1546692994 (large-100-pages / tracked-changes) failed at step 1:",
      "  0. story batch: deleteBlock",
      "  1. story batch: insertAfterBlock",
      message,
    ].join("\n"),
  );
const EMPTIED_JOIN =
  'step 1: the result is not what was asked (insertAfterBlock):\n  "Inserted.": kind is "listItem", expected "paragraph"\n    before undefined\n    after {"kind":"listItem","text":"Inserted."}';

describe("generated flows that reproduce a recorded finding", () => {
  test("the same failure on the canonical session is no finding", () => {
    expect(
      recordedFindingOf(generatedFailure(EMPTIED_JOIN), "canonical"),
    ).toBeUndefined();
  });

  test("a finding of both engines is that finding on the canonical session", () => {
    expect(recordedFindingOf(generatedFailure(KNOWN), "canonical")).toBe(
      FINDING,
    );
  });

  test("a recorded finding logs no failure marker", () => {
    const logged: string[] = [];
    const capture = (...parts: unknown[]) => {
      logged.push(parts.map(String).join(" "));
    };
    const spies = [
      spyOn(console, "error").mockImplementation(capture),
      spyOn(console, "warn").mockImplementation(capture),
      spyOn(console, "log").mockImplementation(capture),
    ];
    try {
      logRecordedFinding("LEGACY_INSERT_AFTER_EMPTIED_JOIN", FLOW);
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
    expect(logged.join("\n")).toContain("LEGACY_INSERT_AFTER_EMPTIED_JOIN");
    expect(extractFailureMarkers(logged.join("\n"))).toEqual([]);
  });
});
