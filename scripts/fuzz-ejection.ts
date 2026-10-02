#!/usr/bin/env bun
/**
 * Tell a merge group's fuzz failure that was already on main from one the
 * group brought in, by replaying the group's failing seeds on main.
 *
 *   replay  In a checkout of the main commit, rerun the consumer scenarios
 *           under the seed the group's run derived from its commit, and each
 *           failed property under its seed and counterexample path; write one
 *           log per replay. Runs main's code, so it holds no write token.
 *   report  Compare the fingerprints the group's logs failed with to the ones
 *           the main replays failed with. Pre-existing ones get their issue
 *           (scripts/fuzz-failure-issues.ts); when every failure was
 *           pre-existing the pull request is labelled for another queue run.
 *           Prints one `FOLIO_QUEUE_EJECTION {json}` line either way.
 *
 * Usage:
 *   bun scripts/fuzz-ejection.ts replay --queue-logs <dir> --group-sha <sha>
 *     --base <checkout> --base-sha <sha> --out <dir>
 *   bun scripts/fuzz-ejection.ts report --queue-logs <dir> --base-logs <dir>
 *     --group-sha <sha> --base-sha <sha> [--pr <n>] [--run-url <url>] [--dry-run]
 */

import { $ } from "bun";
import { TaggedError } from "better-result";
import { spawnSync } from "node:child_process";
import {
  appendFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import { stripVTControlCharacters } from "node:util";

import { hash32 } from "../test/commit-seed";
import { extractFailureMarkers } from "./failure-fingerprints";
import { type Context, type Finding, fileFindings, readFindings } from "./fuzz-failure-issues";
import type { IssueStore } from "./fuzz-issue-classes";

export const LABEL = {
  name: "queue-ejection-pre-existing",
  color: "fbca04",
  description: "Ejected from the merge queue by a failure reproduced on main",
};

const SHA = /^[0-9a-f]{40}$/u;

/** Every `*.log` under `dir`. */
export const logsIn = (dir: string): string[] =>
  existsSync(dir)
    ? readdirSync(dir, { recursive: true })
        .map(String)
        .filter((file) => file.endsWith(".log"))
        .map((file) => path.join(dir, file))
    : [];

/** The fuzz seed scripts/consumer-scenarios.ts derives under CI for commit `sha`. */
export const consumerSeed = (sha: string): number =>
  hash32(`${sha}\0consumer-scenarios fuzz`) >>> 1;

export type PropertyReplay = {
  /** Repo-relative test file. */
  file: string;
  title: string;
  seed: number;
  path: string | null;
  factor: number;
};

/** A `bun test -t` pattern for a title as written: `${…}` matches anything (as in test/property-testing.ts). */
const titlePattern = (title: string): string =>
  title
    .split(/\$\{[^}]*\}/u)
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"))
    .join(".*");

/** The failed properties a log names, from its `PROPERTY_FAILURE {json}` lines. */
export const propertyReplays = (log: string): PropertyReplay[] =>
  log.split("\n").flatMap((line) => {
    const start = line.indexOf("PROPERTY_FAILURE ");
    if (start === -1) return [];
    let value: Record<string, unknown>;
    try {
      value = JSON.parse(line.slice(start + "PROPERTY_FAILURE ".length)) as Record<string, unknown>;
    } catch {
      return [];
    }
    const { file, test, seed, numRunsFactor } = value;
    const counterexample = value["path"];
    if (
      typeof file !== "string" ||
      !/^[\w@.-]+(?:\/[\w@.-]+)*\.test\.tsx?$/u.test(file) ||
      file.split("/").includes("..") ||
      typeof test !== "string" ||
      typeof seed !== "number" ||
      !Number.isSafeInteger(seed) ||
      (counterexample !== null &&
        (typeof counterexample !== "string" || !/^[\d:]*$/u.test(counterexample)))
    ) {
      return [];
    }
    return [
      {
        file,
        title: test,
        seed,
        path: counterexample === "" ? null : counterexample,
        factor: typeof numRunsFactor === "number" && numRunsFactor >= 1 ? numRunsFactor : 1,
      },
    ];
  });

/** The directory to run `bun test` in, and the file relative to it. */
export const packageOf = (file: string): { dir: string; file: string } => {
  const match = /^(packages\/[^/]+)\/(.+)$/u.exec(file);
  return match === null
    ? { dir: ".", file }
    : { dir: match[1] as string, file: match[2] as string };
};

// ---------------------------------------------------------------------------
// replay
// ---------------------------------------------------------------------------

export class ReplayVerificationError extends TaggedError("ReplayVerificationError")<{
  message: string;
}> {}

export type ReplayOutcome =
  | { status: "passed"; fingerprints: string[] }
  | { status: "failed"; fingerprints: string[] }
  | { status: "unavailable"; message: string };
export type ReplayAttempt = { requested: string[]; outcome: ReplayOutcome };

/** A setup failure or a zero-test run is not evidence that main passed. */
export const replayOutcome = ({
  output,
  exitCode,
  error,
}: {
  output: string;
  exitCode: number | null;
  error?: string;
}): ReplayOutcome => {
  const text = stripVTControlCharacters(output);
  const ranTests =
    /^\s*(?:ℹ\s*)?(?:pass|fail)\s+[1-9]\d*\b/mu.test(text) ||
    /^\s*[1-9]\d*\s+(?:pass|fail)\b/mu.test(text);
  if (error || exitCode === null || !ranTests) {
    return { status: "unavailable", message: error ?? "Replay did not execute tests" };
  }
  const fingerprints = extractFailureMarkers(text).map((marker) => marker.fingerprint);
  if (exitCode === 0 && fingerprints.length === 0) return { status: "passed", fingerprints };
  if (exitCode !== 0 && fingerprints.length > 0) return { status: "failed", fingerprints };
  return { status: "unavailable", message: "Replay exit status and failure evidence disagree" };
};

const run = (
  command: string,
  args: string[],
  { cwd, env, log }: { cwd: string; env: Record<string, string>; log: string },
): ReplayOutcome => {
  console.log(`→ ${[command, ...args].join(" ")} (in ${cwd})`);
  const result = spawnSync(command, args, {
    cwd,
    env: { ...process.env, ...env },
    encoding: "utf8",
    maxBuffer: 256 * 1024 * 1024,
    timeout: 45 * 60 * 1_000,
  });
  const output = `${result.stdout ?? ""}\n${result.stderr ?? ""}`;
  appendFileSync(log, output);
  console.log(
    `  exit ${String(result.status)}; ${String(extractFailureMarkers(output).length)} failure line(s)`,
  );
  const outcome = replayOutcome({
    output,
    exitCode: result.status,
    ...(result.error && { error: result.error.message }),
  });
  if (outcome.status === "unavailable")
    throw new ReplayVerificationError({ message: outcome.message });
  return outcome;
};

const replay = (options: Options): void => {
  if (!SHA.test(options.groupSha) || !SHA.test(options.baseSha)) {
    throw new ReplayVerificationError({ message: "Replay requires group and main commit shas" });
  }
  mkdirSync(options.out, { recursive: true });
  const logs = logsIn(options.queueLogs).map((file) => readFileSync(file, "utf8"));
  const markers = logs.flatMap((log) => extractFailureMarkers(log));
  const attempts: ReplayAttempt[] = [];
  const consumers = markers.filter((marker) => marker.repro.includes("consumer-scenarios.ts"));
  if (consumers.length > 0) {
    // The group's run seeded its fuzz flows from the group commit; the base
    // replays the same flows under that seed, with the same defaults.
    const outcome = run("bun", ["scripts/consumer-scenarios.ts"], {
      cwd: options.base,
      env: { FOLIO_SCENARIO_SEED: String(consumerSeed(options.groupSha)) },
      log: path.join(options.out, "base-consumer-scenarios.log"),
    });
    attempts.push({ requested: consumers.map((marker) => marker.fingerprint), outcome });
  }
  const seen = new Set<string>();
  logs.flatMap(propertyReplays).forEach((property, index) => {
    const key = `${property.file}\0${property.title}\0${String(property.seed)}`;
    if (seen.has(key)) return;
    seen.add(key);
    const { dir, file } = packageOf(property.file);
    const outcome = run("bun", ["test", file, "-t", titlePattern(property.title)], {
      cwd: path.join(options.base, dir),
      env: {
        CI: "true",
        PROPERTY_TEST_SEED: String(property.seed),
        PROPERTY_TEST_NUM_RUNS_FACTOR: String(property.factor),
        ...(property.path === null ? {} : { PROPERTY_TEST_PATH: property.path }),
      },
      log: path.join(options.out, `base-property-${String(index)}.log`),
    });
    attempts.push({
      requested: markers
        .filter(
          (marker) =>
            marker.test === `${property.file}::${property.title}` && marker.seed === property.seed,
        )
        .map((marker) => marker.fingerprint),
      outcome,
    });
  });
  verifyReplayAttempts(readFindings(logsIn(options.queueLogs), []), attempts);
  writeFileSync(
    path.join(options.out, "replay-complete.json"),
    JSON.stringify({
      group: options.groupSha,
      main: options.baseSha,
      attempts,
    }),
  );
};

// ---------------------------------------------------------------------------
// report
// ---------------------------------------------------------------------------

export type Verdict = "none" | "pre-existing" | "introduced" | "mixed";

export type Ejection = {
  pr: number | null;
  group: string;
  base: string;
  verdict: Verdict;
  /** Whether the pull request can go back into the queue as it is. */
  requeue: boolean;
  fingerprints: {
    fingerprint: string;
    test: string;
    reproducedOnBase: boolean;
    issue: string | null;
  }[];
};

/** Which of the group's failures the main replays failed with too. */
export const judge = (
  group: readonly { fingerprint: string; test: string }[],
  base: ReadonlySet<string>,
): Pick<Ejection, "verdict" | "requeue" | "fingerprints"> => {
  const fingerprints = group.map(({ fingerprint, test }) => ({
    fingerprint,
    test,
    reproducedOnBase: base.has(fingerprint),
    issue: null as string | null,
  }));
  const reproduced = fingerprints.filter((entry) => entry.reproducedOnBase).length;
  let verdict: Verdict = "mixed";
  if (fingerprints.length === 0) verdict = "none";
  else if (reproduced === fingerprints.length) verdict = "pre-existing";
  else if (reproduced === 0) verdict = "introduced";
  return { verdict, requeue: verdict === "pre-existing", fingerprints };
};

/** Validate every replay before any issue-store access, including mixed failures. */
export const verifyReplayAttempts = (
  findings: readonly Finding[],
  attempts: readonly ReplayAttempt[],
): void => {
  const covered = new Set<string>();
  for (const { requested, outcome } of attempts) {
    if (outcome.status === "unavailable")
      throw new ReplayVerificationError({ message: outcome.message });
    if ((outcome.status === "failed") !== outcome.fingerprints.length > 0) {
      throw new ReplayVerificationError({
        message: "Replay outcome lacks matching execution evidence",
      });
    }
    for (const fingerprint of requested) covered.add(fingerprint);
  }
  if (findings.some(({ record }) => !covered.has(record.marker.fingerprint))) {
    throw new ReplayVerificationError({
      message: "A merge-group finding has no completed main replay",
    });
  }
};

const judgeMainFindings = (findings: readonly Finding[], attempts: readonly ReplayAttempt[]) => {
  verifyReplayAttempts(findings, attempts);
  const reproduced = new Set(
    attempts.flatMap(({ outcome }) => (outcome.status === "failed" ? outcome.fingerprints : [])),
  );
  return judge(
    findings.map(({ record }) => record.marker),
    reproduced,
  );
};

type FileMainFindingsOptions = {
  findings: Finding[];
  attempts: ReplayAttempt[];
  context: Context;
  root: string;
  store?: IssueStore;
};
export const fileMainFindings = async ({
  findings,
  attempts,
  context,
  root,
  store,
}: FileMainFindingsOptions) => {
  const judged = judgeMainFindings(findings, attempts);
  const reproduced = new Set(
    attempts.flatMap(({ outcome }) => (outcome.status === "failed" ? outcome.fingerprints : [])),
  );
  const confirmed = findings.filter(({ record }) => reproduced.has(record.marker.fingerprint));
  if (confirmed.length > 0) {
    const filed = await fileFindings({
      findings: confirmed,
      context,
      root,
      ...(store && { store }),
    });
    for (const entry of judged.fingerprints) {
      entry.issue =
        filed.find(({ fingerprint }) => fingerprint === entry.fingerprint)?.issue ?? null;
    }
  }
  return judged;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const fingerprintsOf = (value: unknown): value is string[] =>
  Array.isArray(value) &&
  value.every((item) => typeof item === "string" && /^[0-9a-f]{16}$/u.test(item));

export const readReplayAttempts = (options: {
  baseLogs: string;
  groupSha: string;
  baseSha: string;
}): ReplayAttempt[] => {
  const file = path.join(options.baseLogs, "replay-complete.json");
  if (!existsSync(file))
    throw new ReplayVerificationError({ message: "Main replay did not complete" });
  const value: unknown = JSON.parse(readFileSync(file, "utf8"));
  if (
    !isRecord(value) ||
    value["group"] !== options.groupSha ||
    value["main"] !== options.baseSha ||
    !Array.isArray(value["attempts"])
  ) {
    throw new ReplayVerificationError({ message: "Main replay evidence does not match this run" });
  }
  return value["attempts"].map((attempt: unknown) => {
    if (
      !isRecord(attempt) ||
      !fingerprintsOf(attempt["requested"]) ||
      !isRecord(attempt["outcome"])
    ) {
      throw new ReplayVerificationError({ message: "Malformed main replay evidence" });
    }
    const { status, fingerprints, message } = attempt["outcome"];
    const requested = attempt["requested"];
    if (status === "unavailable" && typeof message === "string")
      return { requested, outcome: { status, message } };
    if ((status === "passed" || status === "failed") && fingerprintsOf(fingerprints))
      return { requested, outcome: { status, fingerprints } };
    throw new ReplayVerificationError({ message: "Malformed main replay outcome" });
  });
};

const report = async (options: Options): Promise<void> => {
  if (!SHA.test(options.groupSha) || !SHA.test(options.baseSha)) {
    throw new ReplayVerificationError({
      message: "--group-sha and --base-sha must be commit shas",
    });
  }
  const findings = readFindings(logsIn(options.queueLogs), []);
  const attempts = readReplayAttempts(options);
  const context: Context = {
    runUrl: options.runUrl,
    sha: options.baseSha,
    source: "a merge queue run, reproduced on main",
    date: new Date().toISOString().slice(0, 10),
  };
  const judged = options.dryRun
    ? judgeMainFindings(findings, attempts)
    : await fileMainFindings({
        findings,
        attempts,
        context,
        root: path.resolve(import.meta.dir, ".."),
      });
  const ejection: Ejection = {
    pr: options.pr,
    group: options.groupSha,
    base: options.baseSha,
    ...judged,
  };
  if (!options.dryRun && ejection.requeue && options.pr !== null) {
    await $`gh label create ${LABEL.name} --color ${LABEL.color} --description ${LABEL.description} --force`
      .quiet()
      .nothrow();
    await $`gh pr edit ${String(options.pr)} --add-label ${LABEL.name}`.quiet().nothrow();
  }
  console.log(`FOLIO_QUEUE_EJECTION ${JSON.stringify(ejection)}`);
  const summary = process.env["GITHUB_STEP_SUMMARY"];
  if (summary !== undefined && summary !== "") {
    appendFileSync(
      summary,
      [
        `### Merge group failure: ${ejection.verdict}`,
        "",
        ...ejection.fingerprints.map(
          (entry) =>
            `- \`${entry.fingerprint}\` ${entry.test}: ${entry.reproducedOnBase ? "reproduces on main" : `not reproduced on main; attributed to ejected PR ${ejection.pr === null ? "(unknown)" : `#${ejection.pr}`}`}${entry.issue === null ? "" : ` (${entry.issue})`}`,
        ),
        "",
        "```",
        `FOLIO_QUEUE_EJECTION ${JSON.stringify(ejection)}`,
        "```",
        "",
      ].join("\n"),
    );
  }
};

// ---------------------------------------------------------------------------

type Options = {
  command: string;
  queueLogs: string;
  baseLogs: string;
  base: string;
  out: string;
  groupSha: string;
  baseSha: string;
  pr: number | null;
  runUrl: string | null;
  dryRun: boolean;
};

const parseArgs = (argv: readonly string[]): Options => {
  const options: Options = {
    command: argv[0] ?? "",
    queueLogs: "",
    baseLogs: "",
    base: "",
    out: "",
    groupSha: "",
    baseSha: "",
    pr: null,
    runUrl: null,
    dryRun: false,
  };
  for (let index = 1; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? "";
    if (arg === "--queue-logs") options.queueLogs = value();
    else if (arg === "--base-logs") options.baseLogs = value();
    else if (arg === "--base") options.base = value();
    else if (arg === "--out") options.out = value();
    else if (arg === "--group-sha") options.groupSha = value();
    else if (arg === "--base-sha") options.baseSha = value();
    else if (arg === "--pr") {
      const pr = value();
      options.pr = /^\d+$/u.test(pr) ? Number(pr) : null;
    } else if (arg === "--run-url") options.runUrl = value() || null;
    else if (arg === "--dry-run") options.dryRun = true;
    else throw new Error(`Unknown argument ${String(arg)}`);
  }
  return options;
};

if (import.meta.main) {
  const options = parseArgs(process.argv.slice(2));
  if (options.command === "replay") replay(options);
  else if (options.command === "report") await report(options);
  else throw new Error("Usage: bun scripts/fuzz-ejection.ts replay|report …");
}
