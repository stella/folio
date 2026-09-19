import { describe, expect, test } from "bun:test";

import {
  medianMsPerMegabyte,
  type PerformanceObservation,
  performanceFailures,
} from "./lib/corpus-invariants/performance";

const BYTES_PER_MEGABYTE = 1_048_576;

/** An observation costing exactly `rate` milliseconds per megabyte. */
const atRate = (rate: number, megabytes = 1): PerformanceObservation => ({
  bytes: megabytes * BYTES_PER_MEGABYTE,
  parseMs: rate * megabytes,
  peakRssBytes: 0,
});

describe("medianMsPerMegabyte", () => {
  test("reports no measurable census for an empty list", () => {
    expect(medianMsPerMegabyte([])).toBe(0);
  });

  test("excludes a zero-byte file, which has no cost per megabyte", () => {
    const zeroByte = { bytes: 0, parseMs: 9000, peakRssBytes: 0 };
    expect(medianMsPerMegabyte([zeroByte, atRate(10), atRate(20), atRate(30)])).toBe(20);
  });

  test("reports no measurable census when every file is zero-byte", () => {
    expect(medianMsPerMegabyte([{ bytes: 0, parseMs: 5, peakRssBytes: 0 }])).toBe(0);
  });

  test("takes the middle rate of an odd count, whatever order it arrives in", () => {
    expect(medianMsPerMegabyte([atRate(50), atRate(10), atRate(30)])).toBe(30);
  });

  test("averages the two middle rates of an even count", () => {
    expect(medianMsPerMegabyte([atRate(10), atRate(20), atRate(30), atRate(50)])).toBe(25);
  });

  test("prices by megabyte, not by file, so size alone does not move the median", () => {
    expect(medianMsPerMegabyte([atRate(40, 64), atRate(40, 1), atRate(40, 1024)])).toBe(40);
  });
});

const TEN_TIMES_MESSAGE = "parse cost exceeds ten times the corpus median per megabyte";
const HUNDRED_TIMES_MESSAGE = "parse cost exceeds one hundred times the corpus median per megabyte";

const messages = (observation: PerformanceObservation, median: number): string[] =>
  performanceFailures(observation, median).map(({ message }) => message);

describe("performanceFailures", () => {
  test("says nothing about a file that costs what the corpus costs", () => {
    expect(messages(atRate(12), 10)).toEqual([]);
  });

  test("reports the ten-times bucket between the two thresholds", () => {
    expect(messages(atRate(500), 10)).toEqual([TEN_TIMES_MESSAGE]);
  });

  test("reports the hundred-times bucket above it, and only that one", () => {
    expect(messages(atRate(5000), 10)).toEqual([HUNDRED_TIMES_MESSAGE]);
  });

  /** Both thresholds are exclusive: "exceeds" means strictly above. */
  test("stays quiet at exactly ten times the median", () => {
    expect(messages(atRate(100), 10)).toEqual([]);
  });

  test("keeps exactly one hundred times the median in the ten-times bucket", () => {
    expect(messages(atRate(1000), 10)).toEqual([TEN_TIMES_MESSAGE]);
  });

  test("carries no numbers, so every outlier in a bucket is one signature", () => {
    const [slow] = messages(atRate(5000), 10);
    expect(slow).not.toMatch(/\d/u);
  });

  test("declines a verdict when the census has no measurable median", () => {
    expect(messages(atRate(9000), 0)).toEqual([]);
  });

  test("declines a verdict on a zero-byte file, which has no cost per megabyte", () => {
    expect(messages({ bytes: 0, parseMs: 9000, peakRssBytes: 0 }, 10)).toEqual([]);
  });

  test("attributes the failure to the performance invariant", () => {
    expect(performanceFailures(atRate(5000), 10)).toEqual([
      { invariant: "performance", message: HUNDRED_TIMES_MESSAGE, frame: "-" },
    ]);
  });
});
