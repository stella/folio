import { panic } from "better-result";
import type { TaskResult } from "tinybench";

type BenchmarkRowOptions = { name: string; result: TaskResult };

export const toBenchmarkRow = ({ name, result }: BenchmarkRowOptions) => {
  switch (result.state) {
    case "not-started":
    case "started":
    case "aborted":
    case "errored":
      return { benchmark: name, state: result.state };
    case "completed":
    case "aborted-with-statistics": {
      const samples = result.latency.samples;
      if (samples === undefined || samples.length === 0) {
        return panic("Benchmark reporting requires retained latency samples", {
          name,
        });
      }
      const p95Index = Math.max(0, Math.ceil(samples.length * 0.95) - 1);
      const p95 = samples.at(p95Index);
      if (p95 === undefined)
        return panic("Benchmark percentile has no sample", { name });
      return {
        benchmark: name,
        "median ms": round(result.latency.p50),
        "p95 ms": round(p95),
        samples: result.latency.samplesCount,
        state: result.state,
      };
    }
    default: {
      const unhandled: never = result;
      return panic("Unhandled benchmark result", { result: unhandled });
    }
  }
};

const round = (value: number): number => Number(value.toFixed(3));
