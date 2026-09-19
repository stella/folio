import { describe, expect, test } from "bun:test";

import { baselineFromCensus, compareToBaseline } from "./lib/corpus-baseline";
import { CensusBuilder, type CorpusCensus } from "./lib/corpus-census";
import {
  type ExpectedRefusals,
  compareToExpectedRefusals,
  partitionExpectedRefusals,
  refreshedExpectedRefusals,
} from "./lib/corpus-refusals";
import { CORPUS_INVARIANTS, failureFromAssertion, failureSignature } from "./lib/corpus-signature";

const LOCK_DIGEST = "a".repeat(64);

const file = (name: string) => ({ sourceId: "source", path: name, sha256: name.repeat(8) });

const censusWith = (failuresByFile: Record<string, readonly string[]>): CorpusCensus => {
  const builder = new CensusBuilder(LOCK_DIGEST);
  for (const [name, messages] of Object.entries(failuresByFile)) {
    builder.addChecked(
      file(name),
      messages.map((message) => failureFromAssertion(CORPUS_INVARIANTS.parse, message)),
    );
  }
  return builder.build();
};

const signatureOf = (message: string): string =>
  failureSignature(failureFromAssertion(CORPUS_INVARIANTS.parse, message));

const refusalsFor = (messages: readonly string[], files = 1): ExpectedRefusals => ({
  schemaVersion: 1,
  entries: messages.map((message) => ({
    signature: signatureOf(message),
    reason: `refused: ${message}`,
    files,
  })),
});

describe("expected refusals", () => {
  test("an allowlisted signature leaves the defect baseline", () => {
    const census = censusWith({ a: ["nested too deep"], b: ["text changed"] });
    const { defects, refusals } = partitionExpectedRefusals(
      census,
      refusalsFor(["nested too deep"]),
    );

    expect(defects.signatures.map((entry) => entry.message)).toEqual(["text changed"]);
    expect(refusals.map((entry) => entry.message)).toEqual(["nested too deep"]);
    // The defect baseline written from the partition no longer records it, so
    // a later run that still refuses that file is not a "new signature".
    expect(compareToBaseline(baselineFromCensus(defects), defects)).toEqual([]);
  });

  test("a refusal spreading to more files is a violation", () => {
    const census = censusWith({ a: ["nested too deep"], b: ["nested too deep"] });
    const { refusals } = partitionExpectedRefusals(census, refusalsFor(["nested too deep"]));

    const violations = compareToExpectedRefusals(refusalsFor(["nested too deep"]), refusals);
    expect(violations.map((violation) => violation.kind)).toEqual(["more-files"]);
  });

  test("a refusal that stopped reproducing has to be removed", () => {
    const violations = compareToExpectedRefusals(refusalsFor(["nested too deep"]), []);
    expect(violations.map((violation) => violation.kind)).toEqual(["resolved-refusal"]);
  });

  test("fewer refused files must be rewritten down, like the baseline", () => {
    const census = censusWith({ a: ["nested too deep"] });
    const { refusals } = partitionExpectedRefusals(census, refusalsFor(["nested too deep"], 2));

    const violations = compareToExpectedRefusals(refusalsFor(["nested too deep"], 2), refusals);
    expect(violations.map((violation) => violation.kind)).toEqual(["fewer-files"]);
  });

  test("a rewrite refreshes counts and never the reasons or the membership", () => {
    const census = censusWith({ a: ["nested too deep"], b: ["nested too deep"], c: ["unrelated"] });
    const recorded = refusalsFor(["nested too deep"], 1);
    const { refusals } = partitionExpectedRefusals(census, recorded);

    const rewritten = refreshedExpectedRefusals(recorded, refusals);

    expect(rewritten.entries).toEqual([
      {
        signature: signatureOf("nested too deep"),
        reason: "refused: nested too deep",
        files: 2,
      },
    ]);
  });
});
