import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { Result } from "better-result";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { fromMarkdown } from "@stll/folio-core/markdown/fromMarkdown";
import {
  ensureParaIds,
  EnsureParaIdsError,
  ENSURE_PARA_IDS_REASONS,
} from "@stll/folio-core/docx/ensureParaIds";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { runCorpusChecks } from "./lib/corpus-check";
import {
  DECLARED_REFUSAL_INVARIANTS,
  DECLARED_REFUSAL_REASONS,
  type CorpusDeclaredRefusal,
} from "./lib/corpus-invariants/contract";
import { opErrorOutcome } from "./lib/corpus-invariants/op-outcome";
import { CensusBuilder, mergeCensuses, renderCensus, CORPUS_EVIDENCE } from "./lib/corpus-census";
import {
  FamilyCensusBuilder,
  mergeFamilyCensuses,
  renderFamilyCensus,
} from "./lib/corpus-family-census";
import { baselineFromCensus, compareToBaseline } from "./lib/corpus-baseline";
import {
  assertDeclaredRefusalCounts,
  emptyDeclaredRefusalCounts,
  countDeclaredRefusals,
  DeclaredRefusalAccountingError,
} from "./lib/corpus-declared-refusals";

const OPS = Object.values(DECLARED_REFUSAL_INVARIANTS);
const reason = DECLARED_REFUSAL_REASONS.SIGNED_PACKAGE;
const file = (index: number) => ({
  sourceId: "fixture",
  path: `${index}.docx`,
  sha256: String(index),
});
const unsignedPackage = async () =>
  new Uint8Array(await createDocx(fromMarkdown("Paragraph without an id.")));

const refusedCensus = (refusals: readonly CorpusDeclaredRefusal[]) => {
  const builder = new CensusBuilder("lock", "report-only");
  builder.add(file(0), { kind: "complete", failures: [], declaredRefusals: refusals });
  return builder.build();
};

describe("declared operation refusals", () => {
  test("signed packages reach a typed refusal in both operation invariants and the file result", async () => {
    const unsigned = await unsignedPackage();
    const zip = await JSZip.loadAsync(unsigned);
    zip.file("_xmlsignatures/sig1.xml", '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>');
    const signed = await zip.generateAsync({ type: "uint8array" });
    const result = await runCorpusChecks(signed, {
      invariantBudgetMs: 30_000,
      fileBudgetMs: 60_000,
      only: new Set(OPS),
    });
    expect(result.kind).toBe("checked");
    if (result.kind !== "checked") throw new Error("fixture must be checked");
    expect(result.failures).toEqual([]);
    expect(result.declaredRefusals).toEqual(OPS.map((invariant) => ({ invariant, reason })));
    const census = refusedCensus(result.declaredRefusals);
    expect(census.passed).toBe(0);
    expect(census.failedFiles).toBe(0);
    expect(census.declaredRefusedFiles).toBe(1);
    for (const invariant of OPS) expect(census.declaredRefusals[invariant][reason]).toBe(1);
  });

  test("unsigned work throwing identical message text remains a failure", async () => {
    const unsigned = await unsignedPackage();
    const zip = await JSZip.loadAsync(unsigned);
    zip.file("_xmlsignatures/sig1.xml", '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>');
    const signed = await zip.generateAsync({ type: "uint8array" });
    const signedError = await Result.tryPromise({
      try: () => ensureParaIds(signed),
      catch: (cause: unknown) => cause,
    });
    if (signedError.isOk() || !(signedError.error instanceof EnsureParaIdsError))
      throw new Error("signed fixture must refuse");
    const signedMessage = signedError.error.message;
    const result = await Result.tryPromise({
      try: async () => {
        expect((await ensureParaIds(unsigned)).assigned).toBe(1);
        throw new Error(signedMessage);
      },
      catch: (cause: unknown) => cause,
    });
    if (result.isOk()) throw new Error("fixture must throw");
    for (const invariant of OPS) {
      const outcome = opErrorOutcome({ invariant, error: result.error, timings: {} });
      expect(outcome.status).toBe("evaluated");
      if (outcome.status !== "evaluated") throw new Error("message text granted a refusal");
      expect(outcome.failures).toHaveLength(1);
      expect(outcome.failures.at(0)?.message).toBe(`Error: ${signedMessage}`);
      expect(
        opErrorOutcome({
          invariant,
          error: new EnsureParaIdsError({
            reason: ENSURE_PARA_IDS_REASONS.NAMESPACE_INVALID,
            message: signedMessage,
          }),
          timings: {},
        }).status,
      ).toBe("evaluated");
    }
  });

  test(
    "refusal accounting merges, deduplicates per file and never counts a pass",
    async () => {
      await assertProperty(
        fc.asyncProperty(
          fc.array(fc.array(fc.constantFrom(...OPS), { minLength: 1, maxLength: 4 }), {
            minLength: 1,
            maxLength: 15,
          }),
          async (rows) => {
            const whole = new CensusBuilder("lock", "report-only");
            const family = new FamilyCensusBuilder("lock", "report-only");
            const shards = [
              new CensusBuilder("lock", "report-only"),
              new CensusBuilder("lock", "report-only"),
            ];
            const familyShards = [
              new FamilyCensusBuilder("lock", "report-only"),
              new FamilyCensusBuilder("lock", "report-only"),
            ];
            rows.forEach((row, index) => {
              const declaredRefusals = row.map((invariant) => ({ invariant, reason }));
              const result = { kind: "complete", failures: [], declaredRefusals } as const;
              const observed = {
                file: file(index),
                bytes: 1,
                parseMs: 1,
                peakRssBytes: 1,
                referenceMs: 1,
                producer: "fixture",
                failures: [],
                declaredRefusals,
                timings: {},
                evidence: CORPUS_EVIDENCE.gating,
              };
              whole.add(file(index), result);
              family.add(observed);
              // SAFETY: both shard arrays have exactly two entries.
              shards[index % 2]!.add(file(index), result);
              familyShards[index % 2]!.add(observed);
            });
            const census = whole.build();
            expect(mergeCensuses(shards.map((shard) => shard.build()))).toEqual(census);
            expect(mergeFamilyCensuses(familyShards.map((shard) => shard.build()))).toEqual(
              family.build(),
            );
            expect(census.passed).toBe(0);
            expect(census.declaredRefusedFiles).toBe(rows.length);
            for (const invariant of OPS)
              expect(census.declaredRefusals[invariant][reason]).toBe(
                rows.filter((row) => row.includes(invariant)).length,
              );
            expect(renderCensus(census, 1)).toContain("declared refusal");
            expect(
              renderFamilyCensus({
                census: family.build(),
                signaturesPerFamily: 1,
                slowestPerStage: 1,
              }),
            ).toContain("declared refusal");
          },
        ),
        { numRuns: 30 },
      );
    },
    propertyTestTimeout(30_000),
  );

  test("baseline compares refusal counts in both directions", () => {
    const before = refusedCensus([{ invariant: DECLARED_REFUSAL_INVARIANTS.opInverse, reason }]);
    const after = refusedCensus(OPS.map((invariant) => ({ invariant, reason })));
    const baseline = baselineFromCensus(before);
    expect(compareToBaseline(baseline, before)).toEqual([]);
    expect(compareToBaseline(baseline, after).map((violation) => violation.kind)).toEqual([
      "more-files",
    ]);
    expect(
      compareToBaseline(baselineFromCensus(after), before).map((violation) => violation.kind),
    ).toEqual(["fewer-files"]);
    const { declaredRefusals: _oldCounts, ...oldBaseline } = baseline;
    expect(compareToBaseline(oldBaseline, before).at(0)?.detail).toContain("predates");
  });

  test("unknown refusal names and missing accounting fail closed", () => {
    const counts = emptyDeclaredRefusalCounts();
    Reflect.set(counts, "unknown-family", {});
    expect(() => assertDeclaredRefusalCounts(counts)).toThrow(DeclaredRefusalAccountingError);
    const unknownReason = emptyDeclaredRefusalCounts();
    Reflect.set(unknownReason[DECLARED_REFUSAL_INVARIANTS.opInverse], "unknown-reason", 1);
    expect(() => assertDeclaredRefusalCounts(unknownReason)).toThrow(
      DeclaredRefusalAccountingError,
    );
    const missing = emptyDeclaredRefusalCounts();
    Reflect.deleteProperty(missing, DECLARED_REFUSAL_INVARIANTS.opInverse);
    expect(() => assertDeclaredRefusalCounts(missing)).toThrow(DeclaredRefusalAccountingError);
    const invalid = emptyDeclaredRefusalCounts();
    invalid[DECLARED_REFUSAL_INVARIANTS.opInverse][reason] = -1;
    expect(() => assertDeclaredRefusalCounts(invalid)).toThrow(DeclaredRefusalAccountingError);
    const repeated = emptyDeclaredRefusalCounts();
    countDeclaredRefusals(repeated, [
      { invariant: DECLARED_REFUSAL_INVARIANTS.opInverse, reason },
      { invariant: DECLARED_REFUSAL_INVARIANTS.opInverse, reason },
    ]);
    expect(repeated[DECLARED_REFUSAL_INVARIANTS.opInverse][reason]).toBe(1);
  });
});
