import { describe, expect, test } from "bun:test";

import type { CorpusFileId } from "./lib/corpus-census";
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
  producer,
  failures,
  timings,
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
    const builder = new FamilyCensusBuilder("digest");
    builder.add(observed("one", "word/16", [RESERIALIZE_FAILURE]));
    builder.add(observed("two", "libreoffice/7", [RESERIALIZE_FAILURE]));
    builder.add(observed("three", "word/16", []));
    const census = builder.build();

    expect(census.files).toBe(3);
    expect(census.signatures).toHaveLength(1);
    expect(census.signatures.at(0)?.files).toBe(2);
    expect(census.signatures.at(0)?.producers).toEqual({ "word/16": 1, "libreoffice/7": 1 });
    expect(census.producers).toEqual({ "word/16": 2, "libreoffice/7": 1 });
  });

  test("a file failing twice in one family counts as one failed file there", () => {
    const builder = new FamilyCensusBuilder("digest");
    const second = failureFromAssertion(
      EXTENDED_CORPUS_INVARIANTS.reserialize,
      "a full repack already loses package.b",
    );
    builder.add(observed("one", "word/16", [RESERIALIZE_FAILURE, second]));
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
    const builder = new FamilyCensusBuilder("digest");
    builder.add(observed("slow", "word/16", [], { "reserialize.forced-save": 900 }));
    builder.add(observed("fast", "word/16", [], { "reserialize.forced-save": 3 }));
    const slowest = builder.build().slowest["reserialize.forced-save"] ?? [];
    expect(slowest.map((timing) => timing.file.sha256)).toEqual(["slow", "fast"]);
  });
});

describe("mergeFamilyCensuses", () => {
  /**
   * A sharded run must reduce to the census of one unsharded run over the same
   * files, or the ratchet compares against a number no single run produces.
   */
  test("a merge of two shards equals one run over both files", () => {
    const whole = new FamilyCensusBuilder("digest");
    whole.add(observed("one", "word/16", [RESERIALIZE_FAILURE]));
    whole.add(observed("two", "libreoffice/7", [RESERIALIZE_FAILURE, SCHEMA_FAILURE]));

    const left = new FamilyCensusBuilder("digest");
    left.add(observed("one", "word/16", [RESERIALIZE_FAILURE]));
    const right = new FamilyCensusBuilder("digest");
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
    const builder = new FamilyCensusBuilder("digest");
    builder.add(observed("one", "word/16", []));
    builder.add(observed("two", "word/12", []));
    const late = failureFromAssertion(
      EXTENDED_CORPUS_INVARIANTS.performance,
      "parse cost exceeds ten times the corpus median per megabyte",
    );

    const census = censusWithLateFailures(builder.build(), [
      { file: file("one"), producer: "word/16", failures: [late] },
      { file: file("two"), producer: "word/12", failures: [] },
    ]);
    const signature = census.signatures.at(0);
    expect(signature?.family).toBe(CORPUS_INVARIANT_FAMILIES.performance);
    expect(signature?.files).toBe(1);
    expect(signature?.producers).toEqual({ "word/16": 1 });
    expect(census.totals[CORPUS_INVARIANT_FAMILIES.performance]?.failedFiles).toBe(1);
  });

  test("leaves a census with nothing late exactly as it was", () => {
    const builder = new FamilyCensusBuilder("digest");
    builder.add(observed("one", "word/16", [RESERIALIZE_FAILURE]));
    const before = builder.build();
    expect(censusWithLateFailures(before, [])).toEqual(before);
  });
});

describe("describeProducers", () => {
  test("busiest first, then alphabetical, capped", () => {
    expect(describeProducers({ "word/16": 2, "word/12": 5, "libreoffice/7": 2 }, 2)).toBe(
      "word/12 5, libreoffice/7 2",
    );
  });
});
