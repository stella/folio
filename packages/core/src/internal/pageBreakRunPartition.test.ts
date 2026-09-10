import { describe, expect, test } from "bun:test";

import { partitionRunsAtPageBreaks } from "./pageBreakRunPartition";

const alternatingRanges = (count: number) => {
  const runs = [];
  const pageBreaks = [];
  for (let index = 0; index < count; index += 1) {
    const runStart = index * 2;
    runs.push({ id: `run-${String(index)}`, pmStart: runStart, pmEnd: runStart + 1 });
    pageBreaks.push({ id: `break-${String(index)}`, pmStart: runStart + 1, pmEnd: runStart + 2 });
  }
  return { runs, pageBreaks };
};

const observeRunReads = <T>(runs: T[]) => {
  const readsByIndex = new Map<number, number>();
  const observed = new Proxy(runs, {
    get: (target, property, receiver) => {
      if (typeof property === "string") {
        const index = Number(property);
        if (Number.isInteger(index) && index >= 0) {
          readsByIndex.set(index, (readsByIndex.get(index) ?? 0) + 1);
        }
      }
      return Reflect.get(target, property, receiver);
    },
  });
  return { observed, readsByIndex };
};

describe("page-break run partitioning", () => {
  test.each([16_000, 32_000])(
    "visits each run at most once plus one boundary peek across %i alternating pairs",
    (pairCount) => {
      const { runs, pageBreaks } = alternatingRanges(pairCount);
      const { observed, readsByIndex } = observeRunReads(runs);
      const result = partitionRunsAtPageBreaks(observed, pageBreaks);

      expect(result.type).toBe("partitioned");
      if (result.type !== "partitioned") {
        throw new Error("Expected ordered alternating ranges to partition");
      }
      expect([...readsByIndex.values()].reduce((sum, count) => sum + count, 0)).toBe(
        pairCount * 2 - 1,
      );
      expect([...readsByIndex.values()].every((count) => count <= 2)).toBe(true);
      expect(result.partitions).toHaveLength(pairCount);
      expect(
        result.partitions.every(
          ({ before, pageBreak }, index) =>
            before.length === 1 && before[0] === runs[index] && pageBreak === pageBreaks[index],
        ),
      ).toBe(true);
      expect(result.partitions.at(0)?.before.at(0)).toBe(runs.at(0));
      expect(result.partitions.at(-1)?.pageBreak).toBe(pageBreaks.at(-1));
      expect(result.remaining).toHaveLength(0);
    },
  );

  test("reports a run that overlaps a page-break range", () => {
    const overlappingRun = { pmStart: 0, pmEnd: 3 };
    const pageBreak = { pmStart: 1, pmEnd: 2 };

    expect(partitionRunsAtPageBreaks([overlappingRun], [pageBreak])).toEqual({
      type: "overlap",
      pageBreak,
      run: overlappingRun,
    });
  });
});
