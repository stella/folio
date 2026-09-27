/** Seeded public edit operations over the pinned, cache-only DOCX corpus. */
import path from "node:path";
import { createHash } from "node:crypto";

import {
  corpusCacheRoot,
  loadCorpusLock,
  REPOSITORY_ROOT,
  writeJsonFile,
} from "./lib/corpus-manifest.ts";
import { selectTiers, tierScopedLockDigest } from "./lib/corpus-tiers.ts";
import {
  compareBaseline,
  type EditBaseline,
  type EditReport,
} from "./lib/corpus-edit-fuzz-ratchet.ts";
import type { EditFailure, EditWorkerResult } from "../test/corpus-edit-fuzz-contract.ts";

const BASELINE = path.join(REPOSITORY_ROOT, "corpus", "edit-fuzz-baseline.json");
const WORKER = path.join(REPOSITORY_ROOT, "test", "corpus-edit-fuzz-worker.ts");
const DEFAULT_TIMEOUT = 120_000;

const option = (name: string): string | undefined => {
  const index = process.argv.indexOf(name);
  return index < 0 ? undefined : process.argv[index + 1];
};

const positive = (value: string | undefined, fallback: number): number => {
  if (value === undefined) return fallback;
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number <= 0)
    throw new Error(`expected a positive integer, got ${value}`);
  return number;
};

const seedOf = (hash: string): number => Number.parseInt(hash.slice(0, 8), 16);
const messageOf = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

const check = async (reports: readonly EditReport[]): Promise<void> => {
  const baseline = JSON.parse(await Bun.file(BASELINE).text()) as EditBaseline;
  const { introduced, missing, duplicates, unexpectedIncomplete, recoveredIncomplete } =
    compareBaseline(reports, baseline);
  if (
    introduced.length ||
    missing.length ||
    duplicates.length ||
    unexpectedIncomplete.length ||
    recoveredIncomplete.length
  ) {
    const describe = (name: string, values: readonly unknown[]): string =>
      `${name} ${values.length}: ${JSON.stringify(values.slice(0, 5))}`;
    process.stderr.write(
      `edit fuzz baseline mismatch: ${[
        describe("introduced", introduced),
        describe("missing", missing),
        describe("duplicates", duplicates),
        describe("unexpected incomplete", unexpectedIncomplete),
        describe("recovered incomplete", recoveredIncomplete),
      ].join("; ")}\n`,
    );
    process.exitCode = 1;
  }
};

const main = async (): Promise<void> => {
  if (process.argv[2] === "replay") {
    const reportPath = option("--report");
    if (!reportPath) throw new Error("replay requires --report FILE");
    const index = Number(option("--case") ?? "0");
    const report = JSON.parse(await Bun.file(reportPath).text()) as EditReport;
    const selected = report.cases[index];
    if (!selected) throw new Error(`case ${index} does not exist`);
    const lock = selectTiers(await loadCorpusLock(), [1]);
    const source = lock.sources.find(({ id }) => selected.document.startsWith(`${id}/`));
    const relativePath = source && selected.document.slice(source.id.length + 1);
    const file = source?.files.find(
      ({ path: filePath }) =>
        filePath === relativePath && source.id + "/" + filePath === selected.document,
    );
    if (!source || !file || file.sha256 !== selected.sha256)
      throw new Error("replay case does not match the pinned corpus lock");
    const location = path.join(corpusCacheRoot(), "sources", source.id, file.path);
    const child = Bun.spawnSync(
      [
        "bun",
        WORKER,
        "replay",
        location,
        JSON.stringify(selected.operations),
        ...(process.argv.includes("--skip-sdk") ? ["--skip-sdk"] : []),
      ],
      { cwd: REPOSITORY_ROOT, stdout: "pipe", stderr: "pipe", timeout: DEFAULT_TIMEOUT },
    );
    if (child.exitCode !== 0)
      throw new Error(`replay worker failed: ${new TextDecoder().decode(child.stderr)}`);
    const observed = JSON.parse(new TextDecoder().decode(child.stdout)) as EditFailure | null;
    process.stdout.write(`${JSON.stringify(observed, null, 2)}\n`);
    if (observed?.signature !== selected.signature) process.exitCode = 1;
    return;
  }
  if (process.argv[2] === "check") {
    const files = process.argv.slice(3);
    if (files.length === 0) throw new Error("check requires report files");
    const reports = await Promise.all(
      files.map(async (file) => JSON.parse(await Bun.file(file).text()) as EditReport),
    );
    const cases = reports.flatMap(({ cases: reportCases }) => reportCases);
    const counts = new Map<string, number>();
    for (const { class: kind } of cases) counts.set(kind, (counts.get(kind) ?? 0) + 1);
    const lines = [
      "## Corpus edit fuzz",
      `Documents: ${reports.reduce((sum, report) => sum + report.documents, 0)}; parsed: ${reports.reduce((sum, report) => sum + report.parsed, 0)}; attempted edits: ${reports.reduce((sum, report) => sum + report.attempts, 0)}`,
      ...["timeout", "worker"].map(
        (reason) =>
          `- ${reason}: ${reports.reduce((sum, report) => sum + report.incomplete.filter((item) => item.reason === reason).length, 0)}`,
      ),
      ...[...counts]
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([kind, count]) => `- ${kind}: ${count}`),
      `Baseline failure identities: ${(JSON.parse(await Bun.file(BASELINE).text()) as EditBaseline).identities.length}`,
      "### Minimal replay cases",
      ...cases
        .slice(0, 3)
        .map(
          ({ document, sha256, seed, signature, operations, expected, observed }) =>
            `- ${signature} in ${document} (SHA-256 ${sha256}, seed ${seed}); ops ${JSON.stringify(operations)}; expected ${expected}; observed ${observed}`,
        ),
    ];
    const summary = lines.join("\n");
    process.stdout.write(`${summary}\n`);
    if (process.env["GITHUB_STEP_SUMMARY"])
      await Bun.write(process.env["GITHUB_STEP_SUMMARY"], `${summary}\n`);
    await check(reports);
    return;
  }
  if (process.argv[2] !== "run")
    throw new Error(
      "usage: bun scripts/corpus-edit-fuzz.ts run [--sample N] [--out FILE] [--skip-sdk]",
    );

  const sample = option("--sample");
  const limit = sample ? positive(sample, 1) : undefined;
  const shard = option("--shard") ?? null;
  const shardParts = shard?.match(/^(\d+)\/(\d+)$/u);
  if (
    shard &&
    (!shardParts || Number(shardParts[1]) < 1 || Number(shardParts[1]) > Number(shardParts[2]))
  )
    throw new Error("--shard expects k/n with 1 <= k <= n");
  if (shard && limit) throw new Error("--shard and --sample cannot be combined");
  const timeout = positive(option("--timeout"), DEFAULT_TIMEOUT);
  if (!process.argv.includes("--skip-sdk")) {
    const dll = path.join(
      REPOSITORY_ROOT,
      "packages/core/scripts/differential/dotnet/bin/Release/net8.0/OpenXmlProjector.dll",
    );
    if (!Bun.which("dotnet") || !(await Bun.file(dll).exists())) {
      throw new Error(
        "Open XML SDK projector is unavailable; build it with dotnet 8 or use --skip-sdk for a local sample",
      );
    }
  }
  const cache = corpusCacheRoot();
  const fullLock = await loadCorpusLock();
  const lock = selectTiers(fullLock, [1]);
  const seen = new Set<string>();
  const entries = lock.sources
    .flatMap((source) =>
      source.files.map((file) => ({
        id: `${source.id}/${file.path}`,
        path: path.join(cache, "sources", source.id, file.path),
        sha256: file.sha256,
      })),
    )
    .filter(({ sha256 }) => {
      if (seen.has(sha256)) return false;
      seen.add(sha256);
      return true;
    });
  let selected = entries;
  if (limit) {
    selected = entries
      .filter((_, index) => index % Math.ceil(entries.length / limit) === 0)
      .slice(0, limit);
  } else if (shardParts) {
    selected = entries.filter(
      (_, index) => index % Number(shardParts[2]) === Number(shardParts[1]) - 1,
    );
  }
  const report: EditReport = {
    schemaVersion: 1,
    lockDigest: tierScopedLockDigest(fullLock, [1]),
    shard,
    sample: limit ?? null,
    documents: selected.length,
    parsed: 0,
    attempts: 0,
    counts: {},
    cases: [],
    incomplete: [],
  };
  for (const [index, entry] of selected.entries()) {
    if (!(await Bun.file(entry.path).exists()))
      throw new Error(`corpus cache is missing ${entry.id}; run bun run corpus:fetch`);
    const seed = seedOf(entry.sha256);
    // oxlint-disable-next-line no-await-in-loop -- each document has its own child process and deadline
    const child = Bun.spawnSync(
      [
        "bun",
        WORKER,
        "worker",
        entry.path,
        String(seed),
        ...(process.argv.includes("--skip-sdk") ? ["--skip-sdk"] : []),
      ],
      {
        cwd: REPOSITORY_ROOT,
        stdout: "pipe",
        stderr: "pipe",
        timeout,
      },
    );
    let result: EditWorkerResult;
    if (child.exitCode !== 0 || child.exitedDueToTimeout) {
      report.incomplete.push({
        document: entry.id,
        sha256: entry.sha256,
        seed,
        reason: child.exitedDueToTimeout ? "timeout" : "worker",
        detail: child.exitedDueToTimeout
          ? `${timeout}ms deadline`
          : `stderr SHA-256 ${createHash("sha256").update(child.stderr).digest("hex")}`,
      });
      continue;
    } else {
      try {
        result = JSON.parse(new TextDecoder().decode(child.stdout));
      } catch (error) {
        throw new Error(`${entry.id}: invalid worker result: ${messageOf(error)}`);
      }
    }
    if (result.status === "parsed") report.parsed += 1;
    report.attempts += result.attempts;
    for (const failure of result.failures) {
      report.counts[failure.class] = (report.counts[failure.class] ?? 0) + 1;
      report.cases.push({ document: entry.id, sha256: entry.sha256, seed, ...failure });
    }
    if ((index + 1) % 10 === 0 || index + 1 === selected.length) {
      process.stderr.write(
        `edit fuzz: ${index + 1}/${selected.length} documents, ${report.cases.length} findings\n`,
      );
    }
  }
  const out = option("--out");
  if (out) await writeJsonFile(out, report);
  const summary = [
    "## Corpus edit fuzz",
    `Documents: ${report.documents}; parsed: ${report.parsed}; attempted edits: ${report.attempts}`,
    `Incomplete: ${report.incomplete.length} (${report.incomplete.filter(({ reason }) => reason === "timeout").length} timed out)`,
    ...Object.entries(report.counts)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([kind, count]) => `- ${kind}: ${count}`),
    "### Minimal replay cases",
    ...report.cases
      .slice(0, 3)
      .map(
        ({ document, sha256, seed, signature, operations, expected, observed }) =>
          `- ${signature} in ${document} (SHA-256 ${sha256}, seed ${seed}); ops ${JSON.stringify(operations)}; expected ${expected}; observed ${observed.replaceAll("\n", " ")}`,
      ),
  ].join("\n");
  process.stdout.write(`${summary}\n`);
  if (process.env["GITHUB_STEP_SUMMARY"])
    await Bun.write(process.env["GITHUB_STEP_SUMMARY"], `${summary}\n`);
  if (!shard) await check([report]);
};

await main();
