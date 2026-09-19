import { describe, expect, test } from "bun:test";

import { baselineFromCensus, compareToBaseline } from "./lib/corpus-baseline";
import { CensusBuilder, type CorpusCensus } from "./lib/corpus-census";
import { CORPUS_INVARIANTS, failureFromAssertion } from "./lib/corpus-signature";

const LOCK_DIGEST = "a".repeat(64);

const file = (name: string) => ({ sourceId: "source", path: name, sha256: name.repeat(8) });

const censusWith = (failuresByFile: Record<string, readonly string[]>): CorpusCensus => {
  const builder = new CensusBuilder(LOCK_DIGEST);
  for (const [name, messages] of Object.entries(failuresByFile)) {
    builder.addChecked(
      file(name),
      messages.map((message) => failureFromAssertion(CORPUS_INVARIANTS.fixedPoint, message)),
    );
  }
  return builder.build();
};

describe("compareToBaseline", () => {
  test("an unchanged run has no violations", () => {
    const census = censusWith({ a: ["text changed"], b: ["text changed"] });
    expect(compareToBaseline(baselineFromCensus(census), census)).toEqual([]);
  });

  test("a signature the baseline never saw is a violation", () => {
    const baseline = baselineFromCensus(censusWith({ a: ["text changed"] }));
    const violations = compareToBaseline(baseline, censusWith({ a: ["blocks changed"] }));
    expect(violations.map((violation) => violation.kind).toSorted()).toEqual([
      "new-signature",
      "resolved-signature",
    ]);
  });

  test("more files failing the same way is a violation", () => {
    const baseline = baselineFromCensus(censusWith({ a: ["text changed"] }));
    const violations = compareToBaseline(
      baseline,
      censusWith({ a: ["text changed"], b: ["text changed"] }),
    );
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("more-files");
  });

  test("fewer files failing is a violation too: the baseline may only shrink", () => {
    const baseline = baselineFromCensus(censusWith({ a: ["text changed"], b: ["text changed"] }));
    const violations = compareToBaseline(baseline, censusWith({ a: ["text changed"] }));
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("fewer-files");
  });

  test("a baseline entry nothing reproduces must be removed", () => {
    const baseline = baselineFromCensus(censusWith({ a: ["text changed"] }));
    const violations = compareToBaseline(baseline, censusWith({}));
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("resolved-signature");
  });

  test("a baseline measured over a different corpus is refused outright", () => {
    const baseline = {
      ...baselineFromCensus(censusWith({ a: ["x"] })),
      lockDigest: "b".repeat(64),
    };
    const violations = compareToBaseline(baseline, censusWith({ a: ["x"] }));
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("corpus-changed");
  });

  test("the recorded baseline is sorted by signature, so a rewrite is a readable diff", () => {
    const baseline = baselineFromCensus(censusWith({ a: ["zebra"], b: ["alpha"], c: ["alpha"] }));
    expect(baseline.entries.map((entry) => entry.signature)).toEqual(
      baseline.entries.map((entry) => entry.signature).toSorted(),
    );
  });
});
