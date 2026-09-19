import { describe, expect, test } from "bun:test";

import type { CorpusFileId } from "./lib/corpus-census";
import {
  FAMILY_BASELINE_FAMILIES,
  compareFamilyToBaseline,
  familyBaselineFromCensus,
  familyBaselinePath,
} from "./lib/corpus-family-baseline";
import { FamilyCensusBuilder, type FamilyCensus } from "./lib/corpus-family-census";
import {
  CORPUS_INVARIANT_FAMILIES,
  EXTENDED_CORPUS_INVARIANTS,
  EXTENDED_INVARIANT_FAMILY,
} from "./lib/corpus-invariants/contract";
import { type CorpusFailure, failureFromAssertion } from "./lib/corpus-signature";

const file = (name: string): CorpusFileId => ({
  sourceId: "synthetic",
  path: `${name}.docx`,
  sha256: name,
});

const RESERIALIZE = CORPUS_INVARIANT_FAMILIES.reserialize;

const failure = (message: string): CorpusFailure =>
  failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.reserialize, message);

const FIRST = failure("replay hides a serializer difference at package.a");
const SECOND = failure("replay hides a serializer difference at package.b");

const censusOf = (
  digest: string,
  rows: readonly { name: string; failures: readonly CorpusFailure[] }[],
): FamilyCensus => {
  const builder = new FamilyCensusBuilder(digest);
  for (const { name, failures } of rows) {
    builder.add({
      file: file(name),
      bytes: 1024,
      parseMs: 1,
      peakRssBytes: 100,
      producer: "word/16",
      failures,
      timings: {},
    });
  }
  return builder.build();
};

describe("family baseline files", () => {
  test("every extended family owns one, and the core family owns none", () => {
    expect(FAMILY_BASELINE_FAMILIES).toEqual(
      [...new Set(Object.values(EXTENDED_INVARIANT_FAMILY))].sort(),
    );
    expect(FAMILY_BASELINE_FAMILIES).not.toContain(CORPUS_INVARIANT_FAMILIES.core);
  });

  test("a family's file is named after it, under corpus/baselines", () => {
    expect(familyBaselinePath(RESERIALIZE)).toEndWith("corpus/baselines/reserialize.json");
  });

  test("a baseline carries only its own family's signatures", () => {
    const census = censusOf("digest", [
      { name: "one", failures: [FIRST] },
      {
        name: "two",
        failures: [
          failureFromAssertion(
            EXTENDED_CORPUS_INVARIANTS.schemaValidity,
            "word/documentN.xml gained x",
          ),
        ],
      },
    ]);
    const baseline = familyBaselineFromCensus(census, RESERIALIZE);
    expect(baseline.entries.map((entry) => entry.message)).toEqual([FIRST.message]);
    expect(baseline.failedFiles).toBe(1);
  });
});

describe("compareFamilyToBaseline", () => {
  const baseline = familyBaselineFromCensus(
    censusOf("digest", [
      { name: "one", failures: [FIRST] },
      { name: "two", failures: [FIRST] },
    ]),
    RESERIALIZE,
  );

  test("the same findings are no violation", () => {
    const census = censusOf("digest", [
      { name: "one", failures: [FIRST] },
      { name: "two", failures: [FIRST] },
    ]);
    expect(compareFamilyToBaseline(baseline, census)).toEqual([]);
  });

  test("a signature that gained files is a regression", () => {
    const census = censusOf("digest", [
      { name: "one", failures: [FIRST] },
      { name: "two", failures: [FIRST] },
      { name: "three", failures: [FIRST] },
    ]);
    expect(compareFamilyToBaseline(baseline, census).map((v) => v.kind)).toEqual(["more-files"]);
  });

  /** An improvement must be written down, so an accidental re-regression cannot hide in old slack. */
  test("a signature that improved must be rewritten before the gate passes", () => {
    const census = censusOf("digest", [{ name: "one", failures: [FIRST] }]);
    expect(compareFamilyToBaseline(baseline, census).map((v) => v.kind)).toEqual(["fewer-files"]);
  });

  test("a signature the corpus has never seen is a new defect", () => {
    const census = censusOf("digest", [
      { name: "one", failures: [FIRST] },
      { name: "two", failures: [FIRST, SECOND] },
    ]);
    expect(compareFamilyToBaseline(baseline, census).map((v) => v.kind)).toEqual(["new-signature"]);
  });

  test("a signature nothing reproduces any more must be removed", () => {
    const census = censusOf("digest", [
      { name: "one", failures: [] },
      { name: "two", failures: [] },
    ]);
    expect(compareFamilyToBaseline(baseline, census).map((v) => v.kind)).toEqual([
      "resolved-signature",
    ]);
  });

  test("a different corpus is refused rather than compared", () => {
    const census = censusOf("another-digest", [
      { name: "one", failures: [FIRST] },
      { name: "two", failures: [FIRST] },
    ]);
    expect(compareFamilyToBaseline(baseline, census).map((v) => v.kind)).toEqual([
      "corpus-changed",
    ]);
  });

  test("another family's regression is not this family's violation", () => {
    const census = censusOf("digest", [
      { name: "one", failures: [FIRST] },
      {
        name: "two",
        failures: [
          FIRST,
          failureFromAssertion(
            EXTENDED_CORPUS_INVARIANTS.editLocality,
            "an unrelated part changed",
          ),
        ],
      },
    ]);
    expect(compareFamilyToBaseline(baseline, census)).toEqual([]);
  });
});
