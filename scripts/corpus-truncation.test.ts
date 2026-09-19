/**
 * Truncation is missing evidence, not evidence of absence.
 *
 * A per-file budget stops a file's run part-way, and which invariants got to
 * run depends on how busy the machine was. The ratchet is exact in both
 * directions, so a truncated file's gating findings cannot be counted — not
 * even the ones it produced before the budget ran out, since a slower run
 * would have stopped sooner and reported fewer.
 *
 * Which files are allowed to stop is committed data
 * (`corpus/report-only-files.json`), so an unlisted file that stops has moved
 * the compared set by the clock: the run is degraded, however small the share.
 */

import { describe, expect, test } from "bun:test";

import {
  baselineFromCensus,
  compareToBaseline,
  isDegradedRun,
  isFailingViolation,
} from "./lib/corpus-baseline";
import { CensusBuilder, type CorpusCensus, type CorpusFileId } from "./lib/corpus-census";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";
import {
  CORPUS_INVARIANTS,
  type CorpusInvariant,
  failureFromAssertion,
} from "./lib/corpus-signature";

const LOCK_DIGEST = "c".repeat(64);
const REPORT_ONLY_DIGEST = "r".repeat(64);

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
    const builder = new CensusBuilder(LOCK_DIGEST, REPORT_ONLY_DIGEST);
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
    const builder = new CensusBuilder(LOCK_DIGEST, REPORT_ONLY_DIGEST);
    builder.add(file("a"), { kind: "truncated", stage: "reserialize", failures: [] });
    const census = builder.build();
    expect(census.files).toBe(1);
    expect(census.passed).toBe(0);
    expect(census.truncated).toBe(1);
  });
});

describe("an unlisted truncation degrades the run", () => {
  const runTruncating = (total: number, truncated: number): CorpusCensus => {
    const builder = new CensusBuilder(LOCK_DIGEST, REPORT_ONLY_DIGEST);
    for (let index = 0; index < total - truncated; index += 1) {
      builder.add(file(`ok${index}`), { kind: "complete", failures: [] });
    }
    for (let index = 0; index < truncated; index += 1) {
      builder.add(file(`slow${index}`), { kind: "truncated", stage: "reserialize", failures: [] });
    }
    return builder.build();
  };

  test("one file in a thousand is enough: no share of the run is tolerated", () => {
    expect(isDegradedRun(runTruncating(1000, 1))).toBe(true);
    expect(isDegradedRun(runTruncating(1000, 0))).toBe(false);
  });

  test("it reports one degraded outcome and nothing else", () => {
    const violations = compareToBaseline(
      baselineFromCensus(runTruncating(1000, 0)),
      runTruncating(1000, 1),
    );
    expect(violations).toHaveLength(1);
    expect(violations[0]?.kind).toBe("run-degraded");
    expect(isFailingViolation(violations[0] ?? { kind: "x" })).toBe(true);
  });

  test("the message names the file, the stage and the way out", () => {
    const violations = compareToBaseline(
      baselineFromCensus(runTruncating(1000, 0)),
      runTruncating(1000, 1),
    );
    expect(violations[0]?.detail).toContain("source/slow0");
    expect(violations[0]?.detail).toContain("reserialize");
    expect(violations[0]?.detail).toContain("corpus/report-only-files.json");
  });

  test("a run that truncated nothing compares normally", () => {
    const census = runTruncating(1000, 0);
    expect(compareToBaseline(baselineFromCensus(census), census)).toEqual([]);
  });
});
