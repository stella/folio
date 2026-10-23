import { expect, test } from "bun:test";
import { Bench, type TaskResult } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "../benchmarks/config";
import { toBenchmarkRow } from "../benchmarks/report";

test("benchmark report covers every result state and retains percentile samples", async () => {
  const bench = new Bench({
    ...MICRO_BENCH_OPTIONS,
    iterations: 2,
    time: 0,
    warmup: false,
  });
  bench.add("sample", () => Math.sqrt(2));
  const task = bench.tasks.at(0);
  if (!task) throw new Error("Missing benchmark task");
  await bench.run();
  const result = task.result;
  if (result.state !== "completed")
    throw new Error("Benchmark did not complete");
  const states = {
    "not-started": { state: "not-started" },
    started: { state: "started" },
    aborted: { state: "aborted" },
    errored: { state: "errored", error: new Error("Fixture failure") },
    completed: result,
    "aborted-with-statistics": {
      state: "aborted-with-statistics",
      latency: result.latency,
      throughput: result.throughput,
      period: result.period,
      totalTime: result.totalTime,
    },
  } as const satisfies Record<TaskResult["state"], TaskResult>;
  for (const state of Object.values(states)) {
    const row = toBenchmarkRow({ name: "sample", result: state });
    expect(row.state).toBe(state.state);
    if (
      state.state === "completed" ||
      state.state === "aborted-with-statistics"
    ) {
      expect(row.samples).toBe(2);
      expect(row["p95 ms"]).toBe(Number(result.latency.max.toFixed(3)));
      expect(Number.isFinite(row["median ms"])).toBe(true);
    }
  }
});

test("benchmark report refuses statistics without retained samples", async () => {
  const bench = new Bench({ iterations: 2, time: 0, warmup: false });
  bench.add("sample", () => Math.sqrt(2));
  await bench.run();
  const task = bench.tasks.at(0);
  if (!task) throw new Error("Missing benchmark task");
  expect(() => toBenchmarkRow(task)).toThrow(
    "requires retained latency samples",
  );
});
