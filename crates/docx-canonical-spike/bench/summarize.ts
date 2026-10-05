/** Keep raw samples; report p50/p95 only from non-warmup measurements. */
import { readFileSync, writeFileSync } from "node:fs";

const file = process.argv.at(2);
if (!file) throw new TypeError("Provide measurements.jsonl.");
const groups = new Map<
  string,
  {
    samples: number[];
    loads: number[];
    bytes: number[];
    rss: number[];
    heap: number[];
    wasm: number[];
  }
>();
let runnerCount = 0;
const fixtures = new Set<string>();
for (const line of readFileSync(file, "utf8").trim().split("\n")) {
  const entry = JSON.parse(line);
  if (entry.type === "runner") {
    runnerCount += 1;
    if (
      runnerCount !== 1 ||
      typeof entry.sourceHash !== "string" ||
      typeof entry.buildRevision !== "string"
    )
      throw new TypeError("Measurements must contain exactly one identified runner/build.");
  }
  if (entry.type === "fixture-verified") fixtures.add(JSON.stringify([entry.name, entry.pages]));
  if (entry.type !== "sample" || entry.warmup) continue;
  if (runnerCount !== 1) throw new TypeError("Samples lack runner provenance.");
  const key = JSON.stringify([entry.name, entry.pages, entry.arm]);
  let group = groups.get(key);
  if (!group) {
    group = { samples: [], loads: [], bytes: [], rss: [], heap: [], wasm: [] };
    groups.set(key, group);
  }
  const elapsed = entry.sample.elapsedMs ?? entry.sample.samplesMs?.at(0);
  if (typeof elapsed !== "number" || !Number.isFinite(elapsed))
    throw new TypeError("Sample lacks finite latency.");
  group.samples.push(elapsed);
  group.loads.push(entry.load1);
  group.bytes.push(entry.sample.outputBytes);
  if (typeof entry.sample.peakRssBytes === "number") group.rss.push(entry.sample.peakRssBytes);
  if (typeof entry.browserMemory === "number") group.heap.push(entry.browserMemory);
  if (typeof entry.sample.wasmMemoryBytes === "number")
    group.wasm.push(entry.sample.wasmMemoryBytes);
}
const quantile = (values: number[], fraction: number) =>
  values
    .slice()
    .sort((a, b) => a - b)
    .at(Math.ceil(values.length * fraction) - 1);
if (runnerCount !== 1) throw new TypeError("Missing runner provenance.");
for (const fixture of fixtures) {
  const [name, pages] = JSON.parse(fixture);
  for (const arm of ["typescript", "wasm", "native"]) {
    const group = groups.get(JSON.stringify([name, pages, arm]));
    if (!group || group.samples.length !== 50)
      throw new TypeError("Each verified fixture needs 50 retained samples for every arm.");
  }
}
if (groups.size !== fixtures.size * 3) throw new TypeError("Unexpected measurement groups.");
const rows = [...groups].map(([key, group]) => ({
  case: JSON.parse(key),
  count: group.samples.length,
  p50Ms: quantile(group.samples, 0.5),
  p95Ms: quantile(group.samples, 0.95),
  loadMin: Math.min(...group.loads),
  loadMax: Math.max(...group.loads),
  outputBytesMax: Math.max(...group.bytes),
  nativeProcessPeakRssBytes: group.rss.length ? Math.max(...group.rss) : null,
  observedBrowserJsHeapBytes: group.heap.length ? Math.max(...group.heap) : null,
  wasmLinearMemoryHighWaterBytes: group.wasm.length ? Math.max(...group.wasm) : null,
}));
if (rows.length === 0)
  throw new TypeError("No measured samples; validation-only logs are not benchmarks.");
writeFileSync(
  `${file}.summary.json`,
  `${JSON.stringify({ source: file, rows, memoryNote: "Native RSS includes process/runtime; browser JS heap is observed high-water, not process peak RSS; WASM linear memory is allocated capacity." }, null, 2)}\n`,
);
console.log(rows);
