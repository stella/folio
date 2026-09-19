/**
 * The gating / report-only partition.
 *
 * A family whose verdict is a wall-clock budget cannot ratchet: the same file
 * crosses the budget on a loaded machine and clears it on an idle one, so a
 * two-way ratchet fails whichever way the next run differs. These tests pin
 * that a report-only family is measured but never compared, and that a gating
 * family still fails in both directions.
 */

import { describe, expect, test } from "bun:test";

import { baselineFromCensus, compareToBaseline } from "./lib/corpus-baseline";
import { CensusBuilder, type CorpusCensus } from "./lib/corpus-census";
import { familyOf } from "./lib/corpus-family-census";
import { FAMILY_BASELINE_FAMILIES } from "./lib/corpus-family-baseline";
import {
  CORPUS_FAMILY_GATING,
  CORPUS_INVARIANT_FAMILIES,
  EXTENDED_CORPUS_INVARIANTS,
  isGatingFamily,
} from "./lib/corpus-invariants/contract";
import {
  CORPUS_INVARIANTS,
  type CorpusInvariant,
  failureFromAssertion,
} from "./lib/corpus-signature";

const LOCK_DIGEST = "b".repeat(64);
const REPORT_ONLY_DIGEST = "r".repeat(64);

const file = (name: string) => ({ sourceId: "source", path: name, sha256: name.repeat(8) });

const censusOf = (
  entries: ReadonlyArray<{ file: string; invariant: CorpusInvariant; message: string }>,
): CorpusCensus => {
  const builder = new CensusBuilder(LOCK_DIGEST, REPORT_ONLY_DIGEST);
  const byFile = new Map<string, Array<{ invariant: CorpusInvariant; message: string }>>();
  for (const entry of entries) {
    byFile.set(entry.file, [...(byFile.get(entry.file) ?? []), entry]);
  }
  for (const [name, failures] of byFile) {
    builder.addChecked(
      file(name),
      failures.map(({ invariant, message }) => failureFromAssertion(invariant, message)),
    );
  }
  return builder.build();
};

const PERFORMANCE = EXTENDED_CORPUS_INVARIANTS.performance;
const GATING = CORPUS_INVARIANTS.fixedPoint;

describe("the family partition is a total, explicit decision", () => {
  test("every family is classified", () => {
    for (const family of Object.values(CORPUS_INVARIANT_FAMILIES)) {
      expect(CORPUS_FAMILY_GATING[family]).toBeDefined();
    }
  });

  test("performance is the report-only family, core still gates", () => {
    expect(isGatingFamily(CORPUS_INVARIANT_FAMILIES.performance)).toBe(false);
    expect(isGatingFamily(CORPUS_INVARIANT_FAMILIES.core)).toBe(true);
  });

  test("a report-only family owns no baseline file", () => {
    expect(FAMILY_BASELINE_FAMILIES).not.toContain(CORPUS_INVARIANT_FAMILIES.performance);
    expect(FAMILY_BASELINE_FAMILIES).toContain(CORPUS_INVARIANT_FAMILIES.reserialize);
  });

  test("the watchdog expiry is a timing finding, not a completeness verdict", () => {
    expect(familyOf(PERFORMANCE)).toBe(CORPUS_INVARIANT_FAMILIES.performance);
    expect(isGatingFamily(familyOf(PERFORMANCE))).toBe(false);
  });
});

describe("report-only findings never reach the shared baseline", () => {
  test("they are not recorded in it", () => {
    const baseline = baselineFromCensus(
      censusOf([
        { file: "a", invariant: GATING, message: "text changed" },
        { file: "a", invariant: PERFORMANCE, message: "parse exceeded its per-file time budget" },
      ]),
    );
    expect(baseline.entries).toHaveLength(1);
    expect(baseline.entries[0]?.invariant).toBe(GATING);
  });

  test("gaining one is not a violation", () => {
    const baseline = baselineFromCensus(
      censusOf([{ file: "a", invariant: GATING, message: "text changed" }]),
    );
    const violations = compareToBaseline(
      baseline,
      censusOf([
        { file: "a", invariant: GATING, message: "text changed" },
        { file: "b", invariant: PERFORMANCE, message: "parse exceeded its per-file time budget" },
      ]),
    );
    expect(violations).toEqual([]);
  });

  test("losing one is not a violation either", () => {
    const baseline = baselineFromCensus(
      censusOf([
        { file: "a", invariant: GATING, message: "text changed" },
        { file: "b", invariant: PERFORMANCE, message: "parse exceeded its per-file time budget" },
      ]),
    );
    const violations = compareToBaseline(
      baseline,
      censusOf([{ file: "a", invariant: GATING, message: "text changed" }]),
    );
    expect(violations).toEqual([]);
  });

  test("a baseline still carrying them from before the partition is not a violation", () => {
    const baseline = baselineFromCensus(
      censusOf([{ file: "a", invariant: GATING, message: "text changed" }]),
    );
    // What a baseline written before the partition looks like.
    baseline.entries.push({
      signature: "performance | parse exceeded its per-file time budget | -",
      invariant: PERFORMANCE,
      message: "parse exceeded its per-file time budget",
      frame: "-",
      files: 3,
    });
    const violations = compareToBaseline(
      baseline,
      censusOf([{ file: "a", invariant: GATING, message: "text changed" }]),
    );
    expect(violations).toEqual([]);
  });

  test("a gating family still fails in both directions", () => {
    const one = censusOf([{ file: "a", invariant: GATING, message: "text changed" }]);
    const two = censusOf([
      { file: "a", invariant: GATING, message: "text changed" },
      { file: "b", invariant: GATING, message: "text changed" },
    ]);
    expect(compareToBaseline(baselineFromCensus(one), two)[0]?.kind).toBe("more-files");
    expect(compareToBaseline(baselineFromCensus(two), one)[0]?.kind).toBe("fewer-files");
  });
});
