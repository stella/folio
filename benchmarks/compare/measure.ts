/**
 * Sampling policy for the compare benchmark.
 *
 * One configuration is measured in one process (see `run.ts`), so a document
 * class cannot leave a warmed-up JIT or a grown heap behind for the next one.
 * Within the process every sample is a whole comparison from bytes to bytes:
 * medians over at least {@link MINIMUM_ITERATIONS} samples after at least
 * {@link MINIMUM_WARMUPS} discarded ones.
 */

import { heapStats } from "bun:jsc";

export const MINIMUM_WARMUPS = 4;
export const MINIMUM_ITERATIONS = 9;

const quantile = (sorted: readonly number[], fraction: number): number => {
  if (sorted.length === 0) {
    return Number.NaN;
  }
  const position = (sorted.length - 1) * fraction;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  const low = sorted[lower] ?? Number.NaN;
  const high = sorted[upper] ?? low;
  return low + (high - low) * (position - lower);
};

export type Distribution = {
  samples: number;
  min: number;
  median: number;
  p95: number;
};

export const summarize = (values: readonly number[]): Distribution => {
  const sorted = [...values].toSorted((left, right) => left - right);
  return {
    samples: sorted.length,
    min: sorted.at(0) ?? Number.NaN,
    median: quantile(sorted, 0.5),
    p95: quantile(sorted, 0.95),
  };
};

/**
 * Bytes the JavaScript heap grew across one un-timed extra run, with a full
 * collection on both sides. Read it as an order of magnitude that must stay
 * linear in document size, not as an exact allocation count: the collector is
 * the only thing that can answer, and it answers approximately.
 */
export const measureHeapGrowth = async (run: () => Promise<unknown>): Promise<number> => {
  Bun.gc(true);
  const before = heapStats().heapSize;
  await run();
  Bun.gc(true);
  return heapStats().heapSize - before;
};

export type SampleOptions<T> = {
  run: () => Promise<T>;
  warmups?: number;
  iterations?: number;
};

export type SampleResult<T> = {
  /** Wall time of the whole run, per sample. */
  wall: readonly number[];
  /** The last sample's value, for the invariant checks to inspect. */
  last: T;
};

export const sample = async <T>({
  run,
  warmups = MINIMUM_WARMUPS,
  iterations = MINIMUM_ITERATIONS,
}: SampleOptions<T>): Promise<SampleResult<T>> => {
  let last = await run();
  for (let index = 1; index < warmups; index++) {
    last = await run();
  }
  const wall: number[] = [];
  for (let index = 0; index < iterations; index++) {
    const start = performance.now();
    last = await run();
    wall.push(performance.now() - start);
  }
  return { wall, last };
};
