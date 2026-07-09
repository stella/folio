import type { BenchOptions } from "tinybench";

/** Explicit statistical settings shared by every microbenchmark group. */
export const MICRO_BENCH_OPTIONS = {
  iterations: 10,
  retainSamples: true,
  time: 500,
  warmup: true,
  warmupIterations: 3,
  warmupTime: 250,
} as const satisfies BenchOptions;
