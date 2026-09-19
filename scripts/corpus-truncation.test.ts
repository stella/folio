/**
 * Truncation is missing evidence, not evidence of absence.
 *
 * A per-file budget stops a file's run part-way, and which invariants got to
 * run depends on how busy the machine was. The ratchet is exact in both
 * directions, so a truncated file's gating findings cannot be counted — not
 * even the ones it produced before the budget ran out, since a slower run
 * would have stopped sooner and reported fewer.
 */

import { describe, expect, test } from "bun:test";

import {
  baselineFromCensus,
  compareToBaseline,
  isDegradedRun,
  isFailingViolation,
} from "./lib/corpus-baseline";
import {
  CensusBuilder,
  type CorpusCensus,
  type CorpusFileId,
  MAX_TRUNCATED_FRACTION,
} from "./lib/corpus-census";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";
import {
  CORPUS_INVARIANTS,
  type CorpusInvariant,
  failureFromAssertion,
} from "./lib/corpus-signature";

const LOCK_DIGEST = "c".repeat(64);

const file = (name: string): CorpusFileId => ({
  sourceId: "source",
  path: name,
  sha256: name.padEnd(8, "0").repeat(8).slice(0, 64),
});

const GATING = CORPUS_INVARIANTS.fixedPoint;
const PERFORMANCE = EXTENDED_CORPUS_INVARIANTS.performance;

const fail = (invariant: CorpusInvariant, message: string) =>
  failureFromAssertion(invariant, message);

describe("a truncated file contributes no gating evidence", () => {
  const censusWithTruncatedFailure = (): CorpusCensus => {
    const builder = new CensusBuilder(LOCK_DIGEST);
    builder.add(file("a"), { kind: "complete", failures: [fail(GATING, "text changed")] });
    builder.add(file("b"), {
      kind: "truncated",
      stage: "reserialize",
      failures: [
        // Produced before the budget ran out: still not evidence.
        fail(GATING, "text changed"),
        fail(PERFORMANCE, "the file time budget was exhausted before reserialize ran"),
      ],
    });
    return builder.build();
  };

  test("its gating findings do not reach the census", () => {
    const census = censusWithTruncatedFailure();
    const gating = census.signatures.find((signature) => signature.invariant === GATING);
    expect(gating?.files).toBe(1);
    expect(gating?.examples.map((example) => example.path)).toEqual(["a"]);
  });

  test("its timing findings do reach the census", () => {
    const census = censusWithTruncatedFailure();
    const timing = census.signatures.find((signature) => signature.invariant === PERFORMANCE);
    expect(timing).toBeDefined();
    expect(timing?.files).toBe(1);
  });

  test("it is counted and named with the stage it stopped at", () => {
    const census = censusWithTruncatedFailure();
    expect(census.truncated).toBe(1);
    expect(census.truncatedExamples).toEqual([{ file: file("b"), stage: "reserialize" }]);
  });

  test("it is not counted as a file that passed", () => {
    const builder = new CensusBuilder(LOCK_DIGEST);
    builder.add(file("a"), { kind: "truncated", stage: "reserialize", failures: [] });
    const census = builder.build();
    expect(census.files).toBe(1);
    expect(census.passed).toBe(0);
    expect(census.truncated).toBe(1);
  });
});

describe("a truncated run cannot report a shrink", () => {
  // One truncated file has to stay under the degraded-run threshold, so these
  // censuses carry a realistic number of healthy files alongside it.
  const HEALTHY = 400;

  const padded = (): CensusBuilder => {
    const builder = new CensusBuilder(LOCK_DIGEST);
    for (let index = 0; index < HEALTHY; index += 1) {
      builder.add(file(`ok${index}`), { kind: "complete", failures: [] });
    }
    return builder;
  };

  const baselineOf = (...names: string[]) => {
    const builder = padded();
    for (const name of names) {
      builder.add(file(name), { kind: "complete", failures: [fail(GATING, "text changed")] });
    }
    return baselineFromCensus(builder.build());
  };

  const runWithTruncation = (completed: string[], truncatedName: string): CorpusCensus => {
    const builder = padded();
    for (const name of completed) {
      builder.add(file(name), { kind: "complete", failures: [fail(GATING, "text changed")] });
    }
    builder.add(file(truncatedName), {
      kind: "truncated",
      stage: "reserialize",
      failures: [fail(GATING, "text changed")],
    });
    return builder.build();
  };

  test("a signature it could not confirm is kept, not resolved", () => {
    const violations = compareToBaseline(baselineOf("a"), runWithTruncation([], "a"));
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("unobserved-truncated");
    expect(violations.every((violation) => !isFailingViolation(violation))).toBe(true);
  });

  test("a signature that lost files is kept, not ratcheted down", () => {
    const violations = compareToBaseline(baselineOf("a", "b"), runWithTruncation(["a"], "b"));
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("unobserved-truncated");
    expect(isFailingViolation(violations[0] ?? { kind: "x" })).toBe(false);
  });

  test("a genuine shrink with no truncation still ratchets", () => {
    const builder = padded();
    builder.add(file("a"), { kind: "complete", failures: [fail(GATING, "text changed")] });
    const violations = compareToBaseline(baselineOf("a", "b"), builder.build());
    expect(violations[0]?.kind).toBe("fewer-files");
    expect(isFailingViolation(violations[0] ?? { kind: "x" })).toBe(true);
  });

  test("truncation cannot hide a regression: more files still fails", () => {
    const builder = padded();
    for (const name of ["a", "b"]) {
      builder.add(file(name), { kind: "complete", failures: [fail(GATING, "text changed")] });
    }
    builder.add(file("c"), {
      kind: "truncated",
      stage: "reserialize",
      failures: [fail(PERFORMANCE, "the file time budget was exhausted before reserialize ran")],
    });
    const violations = compareToBaseline(baselineOf("a"), builder.build());
    expect(violations.find(isFailingViolation)?.kind).toBe("more-files");
  });
});

describe("a degraded run may not be written down", () => {
  test("the write bar and the compare bar are the same rule", () => {
    const builder = new CensusBuilder(LOCK_DIGEST);
    for (let index = 0; index < 100; index += 1) {
      builder.add(file(`ok${index}`), { kind: "complete", failures: [] });
    }
    for (let index = 0; index < 5; index += 1) {
      builder.add(file(`slow${index}`), { kind: "truncated", stage: "reserialize", failures: [] });
    }
    expect(isDegradedRun(builder.build())).toBe(true);
  });

  test("the handful of files that always stop at a budget is not degraded", () => {
    const builder = new CensusBuilder(LOCK_DIGEST);
    for (let index = 0; index < 1970; index += 1) {
      builder.add(file(`ok${index}`), { kind: "complete", failures: [] });
    }
    for (let index = 0; index < 3; index += 1) {
      builder.add(file(`slow${index}`), { kind: "truncated", stage: "reserialize", failures: [] });
    }
    expect(isDegradedRun(builder.build())).toBe(false);
  });
});

describe("a run that truncated too much is not compared at all", () => {
  const runTruncating = (total: number, truncated: number): CorpusCensus => {
    const builder = new CensusBuilder(LOCK_DIGEST);
    for (let index = 0; index < total - truncated; index += 1) {
      builder.add(file(`ok${index}`), { kind: "complete", failures: [] });
    }
    for (let index = 0; index < truncated; index += 1) {
      builder.add(file(`slow${index}`), {
        kind: "truncated",
        stage: "reserialize",
        failures: [],
      });
    }
    return builder.build();
  };

  test("over the threshold it reports one degraded outcome and nothing else", () => {
    const census = runTruncating(100, Math.ceil(100 * MAX_TRUNCATED_FRACTION) + 1);
    const violations = compareToBaseline(baselineFromCensus(runTruncating(100, 0)), census);
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("run-degraded");
    expect(isFailingViolation(violations[0] ?? { kind: "x" })).toBe(true);
  });

  test("at or under the threshold it compares normally", () => {
    const census = runTruncating(1000, 5);
    expect(compareToBaseline(baselineFromCensus(runTruncating(1000, 0)), census)).toEqual([]);
  });
});
