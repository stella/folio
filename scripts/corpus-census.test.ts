import { describe, expect, test } from "bun:test";

import { CensusBuilder, type CorpusFileId, mergeCensuses } from "./lib/corpus-census";
import {
  CORPUS_INVARIANTS,
  type CorpusFailure,
  failureFromAssertion,
} from "./lib/corpus-signature";

const LOCK_DIGEST = "c".repeat(64);
const REPORT_ONLY_DIGEST = "r".repeat(64);

const file = (index: number): CorpusFileId => ({
  sourceId: "source",
  path: `file-${index}.docx`,
  sha256: String(index).padStart(64, "0"),
});

type Observation =
  | { kind: "duplicate" }
  | { kind: "not-a-docx" }
  | { kind: "checked"; failures: CorpusFailure[] };

const observe = (builder: CensusBuilder, index: number, observation: Observation): void => {
  switch (observation.kind) {
    case "duplicate":
      builder.countDuplicate();
      return;
    case "not-a-docx":
      builder.addNotADocx(file(index), "unreadable-archive");
      return;
    case "checked":
      builder.addChecked(file(index), observation.failures);
      return;
    default: {
      const unreachable: never = observation;
      throw new Error(`Unhandled observation: ${String(unreachable)}`);
    }
  }
};

const failure = (message: string): CorpusFailure =>
  failureFromAssertion(CORPUS_INVARIANTS.fixedPoint, message);

const OBSERVATIONS: Observation[] = [
  { kind: "checked", failures: [] },
  { kind: "checked", failures: [failure("text changed")] },
  { kind: "duplicate" },
  { kind: "not-a-docx" },
  { kind: "checked", failures: [failure("text changed"), failure("blocks changed")] },
  { kind: "checked", failures: [failure("text changed")] },
  { kind: "checked", failures: [] },
  { kind: "checked", failures: [failure("blocks changed")] },
];

const censusOf = (indices: readonly number[]) => {
  const builder = new CensusBuilder(LOCK_DIGEST, REPORT_ONLY_DIGEST);
  for (const index of indices) {
    observe(builder, index, OBSERVATIONS[index] as Observation);
  }
  return builder.build();
};

describe("census", () => {
  test("counts every file exactly once", () => {
    const census = censusOf(OBSERVATIONS.map((_, index) => index));
    expect(census.files).toBe(7);
    expect(census.duplicates).toBe(1);
    expect(census.notADocx).toBe(1);
    expect(census.passed).toBe(2);
    expect(census.failedFiles).toBe(4);
  });

  test("merging shards equals running the whole corpus at once", () => {
    const indices = OBSERVATIONS.map((_, index) => index);
    const shards = [0, 1, 2].map((shard) =>
      censusOf(indices.filter((index) => index % 3 === shard)),
    );
    expect(mergeCensuses(shards)).toEqual(censusOf(indices));
  });

  test("keeps at most three examples per signature", () => {
    const builder = new CensusBuilder(LOCK_DIGEST, REPORT_ONLY_DIGEST);
    for (let index = 0; index < 10; index += 1) {
      builder.addChecked(file(index), [failure("text changed")]);
    }
    const [signature] = builder.build().signatures;
    expect(signature?.files).toBe(10);
    expect(signature?.examples).toHaveLength(3);
  });

  test("orders signatures by how many files they affect", () => {
    const census = censusOf(OBSERVATIONS.map((_, index) => index));
    expect(census.signatures.map((entry) => entry.files)).toEqual([3, 2]);
  });
});
