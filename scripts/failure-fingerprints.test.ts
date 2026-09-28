import { describe, expect, test } from "bun:test";

import { failureMarker } from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  classifyFailureMarkers,
  extractFailureMarkers,
  parseKnownFailures,
} from "./failure-fingerprints";

const marker = (seed: number, failure: Error) =>
  failureMarker({
    test: "comments / suggested",
    seed,
    repro: `FOLIO_SCENARIO_SEED=${seed} bun scripts/consumer-scenarios.ts --only 'comments'`,
    failure,
  });

describe("failure fingerprints", () => {
  test("the same mismatch keeps its fingerprint across seeds and document text", () => {
    const first = marker(
      1250352731,
      new Error(
        'step 8: the result is not what was asked (replaceBlock):\n  no comment {"id":1790619452379,"anchor":"The Buyer pays thirty days."} among []',
      ),
    );
    const second = marker(
      1250359999,
      new Error(
        'step 3: the result is not what was asked (formatRange, replaceBlock):\n  no comment {"id":1790619459911,"anchor":"Le contrat est signé."} among []',
      ),
    );
    expect(first.fingerprint).toBe(second.fingerprint);
    expect(first.assertion).toBe("no comment");
  });

  test("different mismatch kinds have different fingerprints", () => {
    const missingComment = marker(1, new Error("step 1: no comment"));
    const changedReader = marker(
      2,
      new Error("step 1: [readerStability] getChanges reads otherwise after the save"),
    );
    expect(missingComment.fingerprint).not.toBe(changedReader.fingerprint);
  });

  test("a non-error cause does not replace the error message", () => {
    const failure = new Error("step 1: no comment", { cause: "request context" });
    expect(marker(1, failure).assertion).toBe("no comment");
  });

  test("extracts CI-prefixed markers and classifies known fingerprints", () => {
    const first = marker(1, new Error("step 1: no comment"));
    const second = marker(2, new Error("step 2: no comment"));
    const other = marker(3, new Error("block texts differ"));
    const log = [first, second, other]
      .map((entry) => `job 2026-09-28 FOLIO_FAILURE ${JSON.stringify(entry)}`)
      .join("\n");
    const classified = classifyFailureMarkers(extractFailureMarkers(log), [
      { fingerprint: first.fingerprint, issueOrPr: "#1191", firstSeen: "2026-09-28" },
    ]);
    expect(classified).toMatchObject([
      { status: "known", seeds: [1, 2], issueOrPr: "#1191" },
      { status: "new", seeds: [3] },
    ]);
  });

  test("rejects malformed known entries", () => {
    expect(() => parseKnownFailures({ known: [{ fingerprint: "bad" }] })).toThrow(
      "Invalid known failure fingerprint entry",
    );
  });
});
