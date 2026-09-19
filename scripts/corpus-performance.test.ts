import { describe, expect, test } from "bun:test";

import {
  type CorpusCostModel,
  type PerformanceObservation,
  fitCorpusCost,
  performanceFailures,
} from "./lib/corpus-invariants/performance";

const MEGABYTE = 1_048_576;

/** A corpus whose parse cost really is `intercept + slope · bytes`. */
const linearCorpus = (
  intercept: number,
  msPerMegabyte: number,
  sizesInMegabytes: readonly number[],
): PerformanceObservation[] =>
  sizesInMegabytes.map((megabytes) => ({
    bytes: megabytes * MEGABYTE,
    parseMs: intercept + msPerMegabyte * megabytes,
    peakRssBytes: 64 * MEGABYTE + 8 * megabytes * MEGABYTE,
  }));

const SIZES = [0.01, 0.02, 0.05, 0.1, 0.25, 0.5, 1, 2, 4, 8, 16] as const;

const fitOrThrow = (observations: readonly PerformanceObservation[]): CorpusCostModel => {
  const model = fitCorpusCost(observations);
  if (!model) {
    throw new Error("the corpus was expected to admit a fit");
  }
  return model;
};

describe("fitCorpusCost", () => {
  test("recovers the fixed overhead and the per-byte cost of a linear corpus", () => {
    const { parseMs } = fitOrThrow(linearCorpus(40, 300, SIZES));
    expect(parseMs.intercept).toBeCloseTo(40, 6);
    expect(parseMs.slope * MEGABYTE).toBeCloseTo(300, 6);
  });

  test("fits peak resident set separately from time", () => {
    const { peakRssBytes } = fitOrThrow(linearCorpus(40, 300, SIZES));
    expect(peakRssBytes.intercept / MEGABYTE).toBeCloseTo(64, 6);
    expect(peakRssBytes.slope).toBeCloseTo(8, 6);
  });

  test("a minority of outliers does not drag the baseline towards them", () => {
    const clean = linearCorpus(40, 300, SIZES);
    const withCliffs = [
      ...clean,
      { bytes: 4 * MEGABYTE, parseMs: 60_000, peakRssBytes: 1400 * MEGABYTE },
      { bytes: 8 * MEGABYTE, parseMs: 90_000, peakRssBytes: 1800 * MEGABYTE },
    ];
    const { parseMs } = fitOrThrow(withCliffs);
    expect(parseMs.slope * MEGABYTE).toBeCloseTo(300, 4);
  });

  test("declines a fit when no two files differ in size", () => {
    expect(fitCorpusCost([{ bytes: MEGABYTE, parseMs: 300, peakRssBytes: 0 }])).toBeNull();
    expect(
      fitCorpusCost([
        { bytes: MEGABYTE, parseMs: 300, peakRssBytes: 0 },
        { bytes: MEGABYTE, parseMs: 900, peakRssBytes: 0 },
      ]),
    ).toBeNull();
  });

  test("ignores zero-byte files, which say nothing about how cost grows", () => {
    const withEmpty = [
      { bytes: 0, parseMs: 9000, peakRssBytes: 4000 * MEGABYTE },
      ...linearCorpus(40, 300, SIZES),
    ];
    expect(fitOrThrow(withEmpty).parseMs.slope * MEGABYTE).toBeCloseTo(300, 6);
  });
});

const TEN_TIMES_MESSAGE = "parse cost exceeds ten times the corpus baseline for its size";
const HUNDRED_TIMES_MESSAGE =
  "parse cost exceeds one hundred times the corpus baseline for its size";
const RSS_MESSAGE = "peak resident set exceeds ten times the corpus baseline for its size";

const MODEL = fitOrThrow(linearCorpus(40, 300, SIZES));

/** A file of `megabytes` whose parse took `times` its predicted cost. */
const atTimesPredicted = (megabytes: number, times: number): PerformanceObservation => ({
  bytes: megabytes * MEGABYTE,
  parseMs: (40 + 300 * megabytes) * times,
  peakRssBytes: 64 * MEGABYTE + 8 * megabytes * MEGABYTE,
});

const messages = (
  observation: PerformanceObservation,
  model: CorpusCostModel | null = MODEL,
): string[] => performanceFailures(observation, model).map(({ message }) => message);

describe("performanceFailures", () => {
  test("says nothing about a file that costs what the corpus costs at its size", () => {
    expect(messages(atTimesPredicted(4, 1.2))).toEqual([]);
  });

  test("reports the ten-times bucket between the two thresholds", () => {
    expect(messages(atTimesPredicted(4, 50))).toEqual([TEN_TIMES_MESSAGE]);
  });

  test("reports the hundred-times bucket above it, and only that one", () => {
    expect(messages(atTimesPredicted(4, 500))).toEqual([HUNDRED_TIMES_MESSAGE]);
  });

  /** Both thresholds are exclusive: "exceeds" means strictly above. */
  test("stays quiet at exactly ten times the prediction", () => {
    expect(messages(atTimesPredicted(4, 10))).toEqual([]);
  });

  test("keeps exactly one hundred times the prediction in the ten-times bucket", () => {
    expect(messages(atTimesPredicted(4, 100))).toEqual([TEN_TIMES_MESSAGE]);
  });

  /**
   * The defect the old rate metric could not express. A small file pays the
   * fixed overhead of opening a package, so dividing it by a hundredth of a
   * megabyte reports thirty times the corpus rate while nothing is wrong.
   */
  test("does not flag a small file for the fixed cost every parse pays", () => {
    expect(messages(atTimesPredicted(0.01, 1))).toEqual([]);
    expect(messages(atTimesPredicted(0.001, 1))).toEqual([]);
  });

  /**
   * The defect the old rate metric hid. A large file whose rate is a few times
   * the corpus median is far past what its own size predicts.
   */
  test("flags a large file running a path the rest of the corpus avoids", () => {
    expect(messages(atTimesPredicted(16, 40))).toEqual([TEN_TIMES_MESSAGE]);
  });

  test("reports memory amplification even when the parse ran at corpus speed", () => {
    expect(
      messages({ bytes: 4 * MEGABYTE, parseMs: 40 + 300 * 4, peakRssBytes: 1400 * MEGABYTE }),
    ).toEqual([RSS_MESSAGE]);
  });

  test("reports time and memory independently, so a file can be both", () => {
    expect(
      messages({ bytes: 4 * MEGABYTE, parseMs: 400_000, peakRssBytes: 1400 * MEGABYTE }),
    ).toEqual([HUNDRED_TIMES_MESSAGE, RSS_MESSAGE]);
  });

  test("carries no numbers, so every outlier in a bucket is one signature", () => {
    for (const message of messages({
      bytes: 4 * MEGABYTE,
      parseMs: 400_000,
      peakRssBytes: 1400 * MEGABYTE,
    })) {
      expect(message).not.toMatch(/\d/u);
    }
  });

  test("declines a verdict when the census admits no fit", () => {
    expect(messages(atTimesPredicted(4, 500), null)).toEqual([]);
  });

  test("declines a verdict on a zero-byte file, which the model does not cover", () => {
    expect(messages({ bytes: 0, parseMs: 9000, peakRssBytes: 4000 * MEGABYTE })).toEqual([]);
  });

  test("attributes every failure to the performance invariant", () => {
    expect(performanceFailures(atTimesPredicted(4, 500), MODEL)).toEqual([
      { invariant: "performance", message: HUNDRED_TIMES_MESSAGE, frame: "-" },
    ]);
  });
});
