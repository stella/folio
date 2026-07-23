/**
 * Cross-language parse benchmark.
 *
 * Unlike the in-process tinybench suite (folio-only, CodSpeed-gated),
 * this spawns each language's DOCX parser on the same fixtures, has each
 * self-time a parse loop, and tabulates the median per-parse time next to folio.
 * It answers "where does a JS DOCX parser sit beside native (Rust) and
 * interpreted (Python) implementations — and what bar should the Rust rewrite
 * clear?" It is an on-demand context table, NOT a CI regression gate.
 *
 * Run: `bun benchmarks/cross-language/run.ts`
 *
 * External parsers are skipped with a note if their toolchain is missing:
 *   - python-docx  (`pip install python-docx`)
 *   - docx-rs      (`cargo build --release` in ./rust)
 *
 * Outputs differ per library (folio builds a full editable model; python-docx
 * and docx-rs build their own), so this measures parse cost, not equivalence.
 */
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { parseDocx as folioParse } from "@stll/folio-core/docx/parser";

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, "../..");
const ITERATIONS = 20;

type FixtureSize = "small" | "medium" | "large";

const FIXTURES: ReadonlyArray<{ size: FixtureSize; path: string }> = [
  { size: "small", path: "tests/visual/fixtures/docx-editor-demo.docx" },
  { size: "medium", path: "tests/visual/fixtures/sample.docx" },
  { size: "large", path: "tests/visual/fixtures/podily-bps.docx" },
];

function median(xs: number[]): number {
  const sorted = [...xs].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 1) {
    return sorted[mid] ?? 0;
  }
  return ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2;
}

async function benchFolio(absPath: string): Promise<number> {
  const bytes = readFileSync(absPath);
  const samples: number[] = [];
  for (let i = 0; i < ITERATIONS; i += 1) {
    const copy = new Uint8Array(bytes);
    const start = performance.now();
    await folioParse(copy);
    samples.push(performance.now() - start);
  }
  return median(samples);
}

/** Spawn a self-timing program that prints `[{ path, median_ms }]` JSON. */
function benchExternal(label: string, cmd: string, args: string[]): Map<string, number> {
  const out = new Map<string, number>();
  const result = spawnSync(cmd, args, { encoding: "utf-8" });
  if (result.status !== 0) {
    const why = (result.stderr || result.error?.message || "").slice(0, 120).trim();
    console.warn(`  ${label}: skipped (${why})`);
    return out;
  }
  try {
    const parsed = JSON.parse(result.stdout) as Array<{ path: string; median_ms: number }>;
    for (const row of parsed) {
      out.set(resolve(row.path), row.median_ms);
    }
  } catch {
    console.warn(`  ${label}: skipped (unparseable output)`);
  }
  return out;
}

function benchRust(absPaths: string[]): Map<string, number> {
  const rustBin = resolve(HERE, "rust/target/release/docx-parse-bench");
  if (!existsSync(rustBin)) {
    console.warn(
      "  docx-rs (Rust): skipped (run `cargo build --release` in benchmarks/cross-language/rust)",
    );
    return new Map<string, number>();
  }
  return benchExternal("docx-rs (Rust)", rustBin, absPaths);
}

const absPaths = FIXTURES.map((fixture) => resolve(REPO_ROOT, fixture.path));
const python = benchExternal("python-docx", "python3", [resolve(HERE, "parse.py"), ...absPaths]);
const rust = benchRust(absPaths);

const fmt = (n: number | undefined): string => (n === undefined ? "—" : n.toFixed(2));

const rows: Array<Record<string, string>> = [];
for (const fixture of FIXTURES) {
  const abs = resolve(REPO_ROOT, fixture.path);
  rows.push({
    fixture: fixture.size,
    "folio · JS (ms)": fmt(await benchFolio(abs)),
    "python-docx (ms)": fmt(python.get(abs)),
    "docx-rs · Rust (ms)": fmt(rust.get(abs)),
  });
}

console.log("\nparse · median ms per parse (lower is faster; outputs differ per library)\n");
console.table(rows);
