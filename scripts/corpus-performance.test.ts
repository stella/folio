import { describe, expect, test } from "bun:test";

import {
  type CorpusCostModel,
  type PerformanceObservation,
  MAX_REFERENCE_MS,
  fitCorpusCost,
  normalizeForLoad,
  performanceFailures,
} from "./lib/corpus-invariants/performance";

const MEGABYTE = 1_048_576;

/**
 * What the reference package costs on a machine doing nothing else.
 *
 * Every observation below carries it, because a file with no reference reading
 * was never priced against the machine and the family declines to judge one.
 */
const IDLE_REFERENCE_MS = 6;

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
    referenceMs: IDLE_REFERENCE_MS,
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
      {
        bytes: 4 * MEGABYTE,
        parseMs: 60_000,
        peakRssBytes: 1400 * MEGABYTE,
        referenceMs: IDLE_REFERENCE_MS,
      },
      {
        bytes: 8 * MEGABYTE,
        parseMs: 90_000,
        peakRssBytes: 1800 * MEGABYTE,
        referenceMs: IDLE_REFERENCE_MS,
      },
    ];
    const { parseMs } = fitOrThrow(withCliffs);
    expect(parseMs.slope * MEGABYTE).toBeCloseTo(300, 4);
  });

  test("declines a fit when no two files differ in size", () => {
    expect(
      fitCorpusCost([
        { bytes: MEGABYTE, parseMs: 300, peakRssBytes: 0, referenceMs: IDLE_REFERENCE_MS },
      ]),
    ).toBeNull();
    expect(
      fitCorpusCost([
        { bytes: MEGABYTE, parseMs: 300, peakRssBytes: 0, referenceMs: IDLE_REFERENCE_MS },
        { bytes: MEGABYTE, parseMs: 900, peakRssBytes: 0, referenceMs: IDLE_REFERENCE_MS },
      ]),
    ).toBeNull();
  });

  test("ignores zero-byte files, which say nothing about how cost grows", () => {
    const withEmpty = [
      { bytes: 0, parseMs: 9000, peakRssBytes: 4000 * MEGABYTE, referenceMs: IDLE_REFERENCE_MS },
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
  referenceMs: IDLE_REFERENCE_MS,
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
      messages({
        bytes: 4 * MEGABYTE,
        parseMs: 40 + 300 * 4,
        peakRssBytes: 1400 * MEGABYTE,
        referenceMs: IDLE_REFERENCE_MS,
      }),
    ).toEqual([RSS_MESSAGE]);
  });

  test("reports time and memory independently, so a file can be both", () => {
    expect(
      messages({
        bytes: 4 * MEGABYTE,
        parseMs: 400_000,
        peakRssBytes: 1400 * MEGABYTE,
        referenceMs: IDLE_REFERENCE_MS,
      }),
    ).toEqual([HUNDRED_TIMES_MESSAGE, RSS_MESSAGE]);
  });

  test("carries no numbers, so every outlier in a bucket is one signature", () => {
    for (const message of messages({
      bytes: 4 * MEGABYTE,
      parseMs: 400_000,
      peakRssBytes: 1400 * MEGABYTE,
      referenceMs: IDLE_REFERENCE_MS,
    })) {
      expect(message).not.toMatch(/\d/u);
    }
  });

  test("declines a verdict when the census admits no fit", () => {
    expect(messages(atTimesPredicted(4, 500), null)).toEqual([]);
  });

  test("declines a verdict on a zero-byte file, which the model does not cover", () => {
    expect(
      messages({
        bytes: 0,
        parseMs: 9000,
        peakRssBytes: 4000 * MEGABYTE,
        referenceMs: IDLE_REFERENCE_MS,
      }),
    ).toEqual([]);
  });

  test("attributes every failure to the performance invariant", () => {
    expect(performanceFailures(atTimesPredicted(4, 500), MODEL)).toEqual([
      { invariant: "performance", message: HUNDRED_TIMES_MESSAGE, frame: "-" },
    ]);
  });

  test("declines a verdict on a file that was never priced against the machine", () => {
    expect(messages({ ...atTimesPredicted(4, 500), referenceMs: 0 })).toEqual([]);
  });
});

/** The same file, re-timed: what it cost and what the reference cost beside it. */
const retimed = (
  observation: PerformanceObservation,
  parseMs: number,
  referenceMs: number,
): PerformanceObservation => ({
  bytes: observation.bytes,
  parseMs,
  peakRssBytes: observation.peakRssBytes,
  referenceMs,
});

/** The same observations, with the machine `slowdown` times slower throughout. */
const underUniformLoad = (
  observations: readonly PerformanceObservation[],
  slowdown: number,
): PerformanceObservation[] =>
  observations.map((observation) =>
    retimed(observation, observation.parseMs * slowdown, observation.referenceMs * slowdown),
  );

const measuredOrThrow = (
  observations: readonly PerformanceObservation[],
): readonly PerformanceObservation[] => {
  const load = normalizeForLoad(observations);
  if (load.status !== "measured") {
    throw new Error(`expected a measured run, got: ${load.reason}`);
  }
  return load.observations;
};

describe("normalizeForLoad", () => {
  /**
   * A machine that is slower by the same factor all run is already handled by
   * the fit: every file and the baseline move together, so the ratio each file
   * is judged by does not move. Normalisation leaves such a run proportional
   * rather than rescaling it to an absolute idle figure, which it has no way
   * to know, and the verdicts come out the same either way.
   */
  test("a uniformly slower machine reaches the same verdicts", () => {
    const idle = linearCorpus(40, 300, SIZES);
    const loaded = measuredOrThrow(underUniformLoad(idle, 8));
    loaded.forEach((observation, index) => {
      expect(observation.parseMs / (idle[index]?.parseMs ?? 1)).toBeCloseTo(8, 6);
    });

    const withCliff = (observations: readonly PerformanceObservation[]) =>
      observations.flatMap((observation) =>
        performanceFailures(observation, fitOrThrow(observations)).map(({ message }) => message),
      );
    expect(withCliff(loaded)).toEqual(withCliff(idle));
  });

  test("leaves peak resident set alone, which load does not inflate", () => {
    const idle = linearCorpus(40, 300, SIZES);
    const loaded = measuredOrThrow(underUniformLoad(idle, 8));
    loaded.forEach((observation, index) => {
      expect(observation.peakRssBytes).toBe(idle[index]?.peakRssBytes ?? 0);
    });
  });

  /**
   * The flap this guard exists for. Load on a shared machine is not constant
   * across a run that takes minutes, so one file lands in a spike and its
   * milliseconds triple while its code does not change. Without the reference
   * it reads as a defect, and reads as fixed on the next run.
   */
  test("a spike over part of a run does not make those files outliers", () => {
    const idle = linearCorpus(40, 300, SIZES);
    const spiky = idle.map((observation, index) =>
      index % 3 === 0
        ? retimed(observation, observation.parseMs * 30, 30 * IDLE_REFERENCE_MS)
        : observation,
    );

    // Unnormalised, the spike is indistinguishable from a slow parse path.
    const rawModel = fitOrThrow(spiky);
    expect(spiky.flatMap((observation) => performanceFailures(observation, rawModel))).not.toEqual(
      [],
    );

    // Normalised, every file is back at what the corpus costs at its size.
    const normalized = measuredOrThrow(spiky);
    const model = fitOrThrow(normalized);
    expect(normalized.flatMap((observation) => performanceFailures(observation, model))).toEqual(
      [],
    );
  });

  /** And the guard must not hide a real cliff while it removes the load. */
  test("a genuine outlier survives normalisation", () => {
    const withCliff = [
      ...linearCorpus(40, 300, SIZES),
      {
        bytes: 4 * MEGABYTE,
        parseMs: 400_000,
        peakRssBytes: 64 * MEGABYTE,
        referenceMs: IDLE_REFERENCE_MS,
      },
    ];
    const normalized = measuredOrThrow(withCliff);
    const model = fitOrThrow(normalized);
    expect(
      normalized.flatMap((observation) =>
        performanceFailures(observation, model).map(({ message }) => message),
      ),
    ).toEqual([HUNDRED_TIMES_MESSAGE]);
  });

  test("declines the run when the reference blew its own budget", () => {
    const load = normalizeForLoad(underUniformLoad(linearCorpus(40, 300, SIZES), 200));
    expect(load.status).toBe("degraded");
    expect(load.status === "degraded" && load.reason).toContain("reference package");
  });

  test("accepts a run whose reference sits at the budget", () => {
    const observations = linearCorpus(40, 300, SIZES).map((observation) =>
      retimed(observation, observation.parseMs, MAX_REFERENCE_MS),
    );
    expect(normalizeForLoad(observations).status).toBe("measured");
  });

  test("declines a run with no reference readings at all", () => {
    const load = normalizeForLoad(
      linearCorpus(40, 300, SIZES).map((observation) =>
        retimed(observation, observation.parseMs, 0),
      ),
    );
    expect(load.status).toBe("degraded");
    expect(load.status === "degraded" && load.reason).toContain("no reference timing");
  });

  /**
   * The budget is read off the run's median, so a handful of spikes cannot
   * silence a family that is otherwise measuring a healthy machine.
   */
  test("a few spikes do not degrade an otherwise quiet run", () => {
    const observations = linearCorpus(40, 300, SIZES).map((observation, index) =>
      index < 2 ? retimed(observation, observation.parseMs, 5_000) : observation,
    );
    expect(normalizeForLoad(observations).status).toBe("measured");
  });

  test("passes a file with no reading through unscaled, for the verdict to decline", () => {
    const observations = linearCorpus(40, 300, SIZES).map((observation, index) =>
      index === 0 ? retimed(observation, observation.parseMs, 0) : observation,
    );
    expect(measuredOrThrow(observations).at(0)?.parseMs).toBe(observations[0]?.parseMs);
  });
});
