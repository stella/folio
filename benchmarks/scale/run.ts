/**
 * CPU-time scale gate. Fixture construction and reviewer setup happen outside
 * every timed operation; each sample gets fresh mutable state. The output is
 * a JSON artifact so CI retains the measured baseline and a reproducible case.
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import process from "node:process";

import { assessComplexity } from "./complexity";
import { SCALE_SCENARIOS } from "./scenarios";

const SAMPLE_COUNT = 3;
const OUTPUT_PATH = ".cache/performance/scale-complexity.json";
type Operation = () => Promise<unknown> | unknown;
const requested = process.env["FOLIO_SCALE_SCENARIO"];
const scenarios = requested
  ? SCALE_SCENARIOS.filter(({ name }) => name === requested)
  : SCALE_SCENARIOS;

if (scenarios.length === 0) {
  throw new Error(`Unknown FOLIO_SCALE_SCENARIO: ${requested}`);
}

const median = (values: number[]): number => {
  const sorted = values.toSorted((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)]!;
};

const measure = async (prepare: () => Promise<Operation> | Operation): Promise<number> => {
  const operation = await prepare();
  global.gc?.();
  const start = process.cpuUsage();
  await operation();
  const elapsed = process.cpuUsage(start);
  return (elapsed.user + elapsed.system) / 1_000;
};

const results = [];
for (const scenario of scenarios) {
  const cpuMs: number[] = [];
  const samples: number[][] = [];
  for (const size of scenario.sizes) {
    const values: number[] = [];
    for (let sample = 0; sample < SAMPLE_COUNT; sample += 1) {
      values.push(await measure(() => scenario.prepare(size)));
    }
    samples.push(values);
    cpuMs.push(median(values));
  }
  const assessment = assessComplexity(scenario.sizes, [cpuMs[0]!, cpuMs[1]!, cpuMs[2]!]);
  const result = { scenario: scenario.name, samples, ...assessment };
  results.push(result);
  console.log(JSON.stringify(result));
}

const report = {
  schemaVersion: 1,
  runtime: process.version,
  sampleCount: SAMPLE_COUNT,
  metric: "process CPU milliseconds (user + system)",
  results,
};
mkdirSync(dirname(OUTPUT_PATH), { recursive: true });
writeFileSync(OUTPUT_PATH, `${JSON.stringify(report, null, 2)}\n`);

if (results.some(({ status }) => status === "fail")) {
  process.exitCode = 1;
}
