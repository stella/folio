#!/usr/bin/env bun
import { panic } from "better-result";
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { discoverTestSuites, focusedTestFiles, type TestTimings } from "./ci-test-shards";

type TimingArtifact = { runId: string; timings: TestTimings };

const median = (values: number[]) => {
  const sorted = values.toSorted((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length === 0) panic("Cannot aggregate empty timing samples");
  if (sorted.length % 2 === 1) return sorted.at(middle) ?? 0;
  return Math.round(((sorted.at(middle - 1) ?? 0) + (sorted.at(middle) ?? 0)) / 2);
};

export const aggregateTestTimings = (artifacts: TimingArtifact[], files: string[]): string => {
  const runs = new Map<string, Map<string, number>>();
  for (const { runId, timings } of artifacts) {
    const run = runs.get(runId) ?? new Map<string, number>();
    for (const [file, duration] of Object.entries(timings)) run.set(file, duration);
    runs.set(runId, run);
  }
  const allMeasurements = [...runs.values()].flatMap((run) => Array.from(run.values()));
  const fallback = median(allMeasurements);
  const result: TestTimings = {};
  for (const file of [...new Set(files)].toSorted()) {
    const samples = [...runs.values()].flatMap((run) => {
      const duration = run.get(file);
      return duration === undefined ? [] : [duration];
    });
    result[file] = samples.length === 0 ? fallback : median(samples);
  }
  return `${JSON.stringify(result, null, 2)}\n`;
};

if (import.meta.main) {
  const args = process.argv.slice(2);
  const outputIndex = args.indexOf("--output");
  const output = args.at(outputIndex + 1);
  const inputs = args.filter((_, index) => index !== outputIndex && index !== outputIndex + 1);
  if (!output || inputs.length === 0) {
    panic("Usage: bun scripts/aggregate-ci-test-timings.ts --output <file> <artifacts...>");
  }
  const artifacts = inputs.map(
    (input) => JSON.parse(readFileSync(input, "utf8")) as TimingArtifact,
  );
  const suites = discoverTestSuites(path.resolve(import.meta.dir, ".."));
  const files = [...suites.flatMap(({ files: suiteFiles }) => suiteFiles), ...focusedTestFiles];
  writeFileSync(output, aggregateTestTimings(artifacts, files));
}
