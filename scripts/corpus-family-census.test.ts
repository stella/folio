import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";

import { CORPUS_EVIDENCE, type CorpusFileId } from "./lib/corpus-census";
import {
  FamilyCensusBuilder,
  censusWithLateFailures,
  describeProducers,
  familyOf,
  mergeFamilyCensuses,
} from "./lib/corpus-family-census";
import {
  CORPUS_INVARIANT_FAMILIES,
  EXTENDED_CORPUS_INVARIANTS,
} from "./lib/corpus-invariants/contract";
import {
  CORPUS_INVARIANTS,
  type CorpusFailure,
  failureFromAssertion,
} from "./lib/corpus-signature";

setDefaultTimeout(propertyTestTimeout(5_000));

const REPORT_ONLY_DIGEST = "r".repeat(64);

const file = (name: string): CorpusFileId => ({
  sourceId: "synthetic",
  path: `${name}.docx`,
  sha256: name,
});

const observed = (
  name: string,
  producer: string,
  failures: readonly CorpusFailure[],
  timings: Record<string, number> = {},
) => ({
  file: file(name),
  bytes: 1024,
  parseMs: 1,
  peakRssBytes: 100,
  referenceMs: 1,
  producer,
  failures,
  declaredRefusals: [],
  timings,
  evidence: CORPUS_EVIDENCE.gating,
});

const RESERIALIZE_FAILURE = failureFromAssertion(
  EXTENDED_CORPUS_INVARIANTS.reserialize,
  "replay hides a serializer difference at package.a",
);
const SCHEMA_FAILURE = failureFromAssertion(
  EXTENDED_CORPUS_INVARIANTS.schemaValidity,
  "word/documentN.xml gained unknown-element",
);

describe("familyOf", () => {
  test("routes every extended invariant to its own family", () => {
    expect(familyOf(EXTENDED_CORPUS_INVARIANTS.reserialize)).toBe(
      CORPUS_INVARIANT_FAMILIES.reserialize,
    );
    expect(familyOf(EXTENDED_CORPUS_INVARIANTS.performance)).toBe(
      CORPUS_INVARIANT_FAMILIES.performance,
    );
  });

  test("leaves the gate's original invariants with the core baseline", () => {
    expect(familyOf(CORPUS_INVARIANTS.parse)).toBe(CORPUS_INVARIANT_FAMILIES.core);
    expect(familyOf(CORPUS_INVARIANTS.completes)).toBe(CORPUS_INVARIANT_FAMILIES.core);
  });
});

describe("FamilyCensusBuilder", () => {
  test("counts a signature once per file and names its producers", () => {
    const builder = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    builder.add(observed("one", "p1/16", [RESERIALIZE_FAILURE]));
    builder.add(observed("two", "libreoffice/7", [RESERIALIZE_FAILURE]));
    builder.add(observed("three", "p1/16", []));
    const census = builder.build();

    expect(census.files).toBe(3);
    expect(census.signatures).toHaveLength(1);
    expect(census.signatures.at(0)?.files).toBe(2);
    expect(census.signatures.at(0)?.producers).toEqual({ "p1/16": 1, "libreoffice/7": 1 });
    expect(census.producers).toEqual({ "p1/16": 2, "libreoffice/7": 1 });
  });

  test("a file failing twice in one family counts as one failed file there", () => {
    const builder = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    const second = failureFromAssertion(
      EXTENDED_CORPUS_INVARIANTS.reserialize,
      "a full repack already loses package.b",
    );
    builder.add(observed("one", "p1/16", [RESERIALIZE_FAILURE, second]));
    const census = builder.build();
    expect(census.totals[CORPUS_INVARIANT_FAMILIES.reserialize]).toEqual({
      files: 1,
      failedFiles: 1,
    });
    expect(census.totals[CORPUS_INVARIANT_FAMILIES.schemaValidity]).toEqual({
      files: 1,
      failedFiles: 0,
    });
  });

  test("keeps the slowest files per stage", () => {
    const builder = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    builder.add(observed("slow", "p1/16", [], { "reserialize.forced-save": 900 }));
    builder.add(observed("fast", "p1/16", [], { "reserialize.forced-save": 3 }));
    const slowest = builder.build().slowest["reserialize.forced-save"] ?? [];
    expect(slowest.map((timing) => timing.file.sha256)).toEqual(["slow", "fast"]);
  });
});

describe("mergeFamilyCensuses", () => {
  test("interleaved shards preserve every file cost in canonical order", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 3, max: 20 }),
        fc.integer({ min: 2, max: 5 }),
        async (count, shardCount) => {
          const whole = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
          const shards = Array.from(
            { length: shardCount },
            () => new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST),
          );
          for (let index = 0; index < count; index += 1) {
            const row = observed(String(index), "fixture", []);
            whole.add(row);
            // SAFETY: modulus is bounded by the constructed shard count.
            shards[index % shardCount]!.add(row);
          }
          const merged = mergeFamilyCensuses(shards.map((shard) => shard.build()));
          expect(merged.costs).toEqual(whole.build().costs);
          expect(merged.costs).toHaveLength(count);
        },
      ),
      { numRuns: 30 },
    );
  });

  /**
   * A sharded run must reduce to the census of one unsharded run over the same
   * files, or the ratchet compares against a number no single run produces.
   */
  test("a merge of two shards equals one run over both files", () => {
    const whole = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    whole.add(observed("one", "p1/16", [RESERIALIZE_FAILURE]));
    whole.add(observed("two", "libreoffice/7", [RESERIALIZE_FAILURE, SCHEMA_FAILURE]));

    const left = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    left.add(observed("one", "p1/16", [RESERIALIZE_FAILURE]));
    const right = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    right.add(observed("two", "libreoffice/7", [RESERIALIZE_FAILURE, SCHEMA_FAILURE]));

    const merged = mergeFamilyCensuses([left.build(), right.build()]);
    expect(merged.files).toBe(whole.build().files);
    expect(merged.producers).toEqual(whole.build().producers);
    expect(
      merged.signatures.map(({ signature, files, producers }) => ({
        signature,
        files,
        producers,
      })),
    ).toEqual(
      whole.build().signatures.map(({ signature, files, producers }) => ({
        signature,
        files,
        producers,
      })),
    );
  });

  test("refuses an empty merge rather than inventing a digest", () => {
    expect(() => mergeFamilyCensuses([])).toThrow();
  });
});

describe("censusWithLateFailures", () => {
  test("folds a verdict taken after the run into the signatures and totals", () => {
    const builder = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    builder.add(observed("one", "p1/16", []));
    builder.add(observed("two", "p1/12", []));
    const late = failureFromAssertion(
      EXTENDED_CORPUS_INVARIANTS.performance,
      "parse cost exceeds ten times the corpus median per megabyte",
    );

    const census = censusWithLateFailures(builder.build(), [
      { file: file("one"), producer: "p1/16", failures: [late] },
      { file: file("two"), producer: "p1/12", failures: [] },
    ]);
    const signature = census.signatures.at(0);
    expect(signature?.family).toBe(CORPUS_INVARIANT_FAMILIES.performance);
    expect(signature?.files).toBe(1);
    expect(signature?.producers).toEqual({ "p1/16": 1 });
    expect(census.totals[CORPUS_INVARIANT_FAMILIES.performance]?.failedFiles).toBe(1);
  });

  test("leaves a census with nothing late exactly as it was", () => {
    const builder = new FamilyCensusBuilder("digest", REPORT_ONLY_DIGEST);
    builder.add(observed("one", "p1/16", [RESERIALIZE_FAILURE]));
    const before = builder.build();
    expect(censusWithLateFailures(before, [])).toEqual(before);
  });
});

describe("describeProducers", () => {
  test("busiest first, then alphabetical, capped", () => {
    expect(describeProducers({ "p1/16": 2, "p1/12": 5, "libreoffice/7": 2 }, 2)).toBe(
      "p1/12 5, libreoffice/7 2",
    );
  });
});
