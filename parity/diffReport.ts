#!/usr/bin/env bun
/**
 * Compare two parity JSON reports and print deterministic score/count deltas.
 */
import path from "node:path";

import { isReferenceRendererId } from "./referenceRenderer";
import type { CorpusReport, DivergenceKind, FeatureAttributedResult } from "./types";

type DiffFlags = {
  before?: string | undefined;
  after?: string | undefined;
};

type DocSummary = {
  file: string;
  score: number;
  total: number;
  counts: Partial<Record<DivergenceKind, number>>;
};

const KIND_ORDER: DivergenceKind[] = [
  "page-count",
  "pagination",
  "missing-line",
  "extra-line",
  "line-break",
  "text-mismatch",
  "y-drift",
  "x-drift",
  "width-drift",
];

const usage = (): string =>
  [
    "Usage:",
    "  bun parity/diffReport.ts before.json after.json",
    "  bun parity/diffReport.ts --before before.json --after after.json",
  ].join("\n");

const parseArgs = (argv: string[]): DiffFlags => {
  const flags: DiffFlags = {};
  const positional: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === undefined) {
      continue;
    }
    if (arg === "--before") {
      flags.before = argv[++i];
    } else if (arg === "--after") {
      flags.after = argv[++i];
    } else if (arg === "--help" || arg === "-h") {
      console.log(usage());
      process.exit(0);
    } else {
      positional.push(arg);
    }
  }
  flags.before ??= positional[0];
  flags.after ??= positional[1];
  if (!flags.before || !flags.after) {
    throw new Error("before and after report paths are required");
  }
  return flags;
};

const readReport = async (file: string): Promise<CorpusReport> => {
  const report = (await Bun.file(path.resolve(file)).json()) as CorpusReport;
  if (!Array.isArray(report.results)) {
    throw new Error(`${file} is not a parity JSON report`);
  }
  if (!report.reference || !isReferenceRendererId(report.reference.id)) {
    throw new Error(`${file} predates explicit reference-renderer metadata; regenerate it`);
  }
  return report;
};

const summarize = (result: FeatureAttributedResult): DocSummary => {
  const counts: Partial<Record<DivergenceKind, number>> = {};
  for (const divergence of result.divergences) {
    counts[divergence.kind] = (counts[divergence.kind] ?? 0) + 1;
  }
  return {
    file: result.file,
    score: result.score,
    total: result.divergences.length,
    counts,
  };
};

const delta = (after: number, before: number): string => {
  const value = after - before;
  if (value > 0) return `+${value}`;
  return String(value);
};

const formatScoreDelta = (after: number, before: number): string => {
  const value = after - before;
  const fixed = value.toFixed(6);
  return value > 0 ? `+${fixed}` : fixed;
};

const main = async (): Promise<void> => {
  const flags = parseArgs(process.argv.slice(2));
  const [before, after] = await Promise.all([readReport(flags.before!), readReport(flags.after!)]);
  if (before.reference.id !== after.reference.id) {
    throw new Error(
      `cannot diff reports from different references (${before.reference.id} vs ${after.reference.id})`,
    );
  }
  const beforeByFile = new Map(before.results.map((result) => [result.file, summarize(result)]));
  const afterByFile = new Map(after.results.map((result) => [result.file, summarize(result)]));
  const files = Array.from(new Set([...beforeByFile.keys(), ...afterByFile.keys()])).toSorted();

  const docs = files.map((file) => {
    const b = beforeByFile.get(file);
    const a = afterByFile.get(file);
    return {
      file,
      before: b,
      after: a,
      scoreDelta: b && a ? formatScoreDelta(a.score, b.score) : null,
      totalDelta: b && a ? delta(a.total, b.total) : null,
      counts: Object.fromEntries(
        KIND_ORDER.map((kind) => {
          const beforeCount = b?.counts[kind] ?? 0;
          const afterCount = a?.counts[kind] ?? 0;
          return [
            kind,
            {
              before: beforeCount,
              after: afterCount,
              delta: delta(afterCount, beforeCount),
            },
          ];
        }),
      ),
    };
  });

  console.log(
    JSON.stringify(
      {
        before: path.resolve(flags.before!),
        after: path.resolve(flags.after!),
        reference: after.reference,
        docs,
      },
      null,
      2,
    ),
  );
};

try {
  await main();
} catch (error) {
  const err = error instanceof Error ? error : new Error(String(error));
  console.error(`parity diff failed: ${err.message}`);
  console.error(usage());
  process.exitCode = 2;
}
