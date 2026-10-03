import { describe, expect, test } from "bun:test";

import { baselineFromCensus } from "./lib/corpus-baseline";
import { CensusBuilder, CORPUS_EVIDENCE } from "./lib/corpus-census";
import {
  compareZeroFamilies,
  compareFamilyToBaseline,
  FAMILY_BASELINE_FAMILIES,
  familyBaselineFromCensus,
  writeFamilyBaselines,
} from "./lib/corpus-family-baseline";
import { FamilyCensusBuilder } from "./lib/corpus-family-census";
import {
  CORPUS_INVARIANT_FAMILIES,
  EXTENDED_CORPUS_INVARIANTS,
  familyOf,
  isGatingFamily,
  isZeroFamily,
} from "./lib/corpus-invariants/contract";
import { failureFromAssertion } from "./lib/corpus-signature";

const FILE = { sourceId: "fixture", path: "document.docx", sha256: "a".repeat(64) };
const OPS = [EXTENDED_CORPUS_INVARIANTS.opInverse, EXTENDED_CORPUS_INVARIANTS.opLocality];

const observed = (invariant: (typeof OPS)[number], evidence: "gating" | "report-only") => {
  const failure = failureFromAssertion(invariant, "generated operation changed an untouched block");
  const builder = new FamilyCensusBuilder("lock", "report-only");
  builder.add({
    file: FILE,
    bytes: 100,
    parseMs: 1,
    peakRssBytes: 100,
    referenceMs: 1,
    producer: "fixture",
    failures: [failure],
    declaredRefusals: [],
    timings: {},
    evidence,
  });
  return { failure, family: builder.build() };
};

describe("operation invariants have an immutable zero allowance", () => {
  for (const invariant of OPS) {
    test(`${invariant} owns a zero family without a writable baseline`, async () => {
      const family = familyOf(invariant);
      expect(isGatingFamily(family)).toBe(true);
      expect(isZeroFamily(family)).toBe(true);
      expect(FAMILY_BASELINE_FAMILIES).not.toContain(family);
      const observation = observed(invariant, CORPUS_EVIDENCE.gating);
      expect(compareZeroFamilies(observation.family)).toHaveLength(1);
      const fabricated = {
        schemaVersion: 1,
        family,
        lockDigest: "lock",
        reportOnlyDigest: "report-only",
        files: 1,
        failedFiles: 1,
        entries: observation.family.signatures,
      } as const;
      expect(compareFamilyToBaseline(fabricated, observation.family)).toHaveLength(1);
      expect(() => familyBaselineFromCensus(observation.family, family)).toThrow("fixed zero");
      // This refuses before any family file can be written.
      await expect(writeFamilyBaselines(observation.family)).rejects.toThrow("cannot be written");
    });

    test(`${invariant} findings survive report-only and truncated evidence`, () => {
      const observation = observed(invariant, CORPUS_EVIDENCE.reportOnly);
      expect(compareZeroFamilies(observation.family)).toHaveLength(1);
      for (const kind of ["report-only", "truncated"] as const) {
        const builder = new CensusBuilder("lock", "report-only");
        builder.add(
          FILE,
          kind === "truncated"
            ? {
                kind,
                declaredRefusals: [],
                failures: [observation.failure],
                stage: "schema-validity",
              }
            : { kind, declaredRefusals: [], failures: [observation.failure] },
        );
        const census = builder.build();
        expect(census.signatures.map((signature) => signature.invariant)).toEqual([invariant]);
        expect(baselineFromCensus(census).entries).toEqual([]);
      }
    });
  }

  test("both families are distinct and an empty census passes", () => {
    expect(familyOf(EXTENDED_CORPUS_INVARIANTS.opInverse)).toBe(
      CORPUS_INVARIANT_FAMILIES.opInverse,
    );
    expect(familyOf(EXTENDED_CORPUS_INVARIANTS.opLocality)).toBe(
      CORPUS_INVARIANT_FAMILIES.opLocality,
    );
    expect(compareZeroFamilies({ signatures: [] })).toEqual([]);
  });
});
