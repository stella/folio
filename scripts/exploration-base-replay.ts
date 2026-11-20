#!/usr/bin/env bun
/**
 * Fail a change whose generated property failure its base commit does not
 * have.
 *
 * Pull request and merge group CI report a generated case that fails in a
 * property file the change does not own instead of failing on it
 * (test/property-exploration.ts): random search there finds failures that
 * were already on main. A change can still break such a file outright, and
 * then every seed reports and nothing blocks. This replays each reported
 * failure's seed and counterexample path, alone, in a checkout of the base
 * commit:
 *
 * - the base fails it too: the failure was already there; it is filed as
 *   every exploration finding is (scripts/fuzz-failure-issues.ts);
 * - the base passes it: the change brought it, and the run fails;
 * - the replay cannot run on the base (no replay line, a stale path, nothing
 *   executed): nothing shows the failure was already there, and the run
 *   fails.
 *
 * Usage:
 *   bun scripts/exploration-base-replay.ts --records <dir> --base <checkout>
 */

import { panic } from "better-result";
import { spawnSync } from "node:child_process";
import { appendFileSync, existsSync, readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import type { FailureRecord } from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  EXPLORATION_BASE_ENV,
  EXPLORATION_DIR_ENV,
  EXPLORATION_ENV,
  EXPLORATION_OWNED_ENV,
} from "../test/property-exploration";
import { NUM_RUNS_FACTOR_ENV } from "../test/property-run-factor";
import {
  CASE_ENV,
  parseReplayCommand,
  PATH_ENV,
  type ReplayCommand,
  SEED_ENV,
} from "../test/replay-command";
import { type ReplayOutcome, replayOutcome } from "./fuzz-ejection";
import { parseFailureRecord } from "./fuzz-failure-issues";

/** The environment a CI test step sets that a base replay must not inherit. */
const STEP_ENV = [
  EXPLORATION_ENV,
  EXPLORATION_DIR_ENV,
  EXPLORATION_BASE_ENV,
  EXPLORATION_OWNED_ENV,
  "FOLIO_PROPERTY_FAILURES_DIR",
  SEED_ENV,
  PATH_ENV,
  CASE_ENV,
  NUM_RUNS_FACTOR_ENV,
] as const;

/** A generated property failure the change under CI reported. */
export type ReportedFailure = {
  test: string;
  fingerprint: string;
  repro: string;
};

/** The exploration findings of TypeScript property files under `dir`. */
export const reportedFailures = (dir: string): ReportedFailure[] => {
  if (!existsSync(dir)) return [];
  const seen = new Set<string>();
  const failures: ReportedFailure[] = [];
  for (const entry of readdirSync(dir, { recursive: true })) {
    const file = String(entry);
    if (!file.endsWith(".json") || file.endsWith(".flow.json")) continue;
    const record: FailureRecord | null = parseFailureRecord(
      JSON.parse(readFileSync(path.join(dir, file), "utf8")),
    );
    if (record === null)
      panic(`${file} is not a failure record; the gate cannot read it.`);
    const { marker } = record;
    if (marker.exploration !== true) continue;
    // Rust property findings are filed by their own lane.
    if (!/\.test\.tsx?$/u.test(marker.test.split("::")[0] ?? "")) continue;
    if (seen.has(marker.repro)) continue;
    seen.add(marker.repro);
    failures.push({
      test: marker.test,
      fingerprint: marker.fingerprint,
      repro: marker.repro,
    });
  }
  return failures;
};

export const VERDICTS = {
  PRE_EXISTING: "preExisting",
  INTRODUCED: "introduced",
  UNREPLAYABLE: "unreplayable",
} as const;

export type Verdict =
  | { type: typeof VERDICTS.PRE_EXISTING }
  | { type: typeof VERDICTS.INTRODUCED }
  | { type: typeof VERDICTS.UNREPLAYABLE; reason: string };

/** What the base's replay says about a failure the change reported. */
export const verdictOf = (outcome: ReplayOutcome): Verdict => {
  switch (outcome.status) {
    case "failed":
      return { type: VERDICTS.PRE_EXISTING };
    case "passed":
      return { type: VERDICTS.INTRODUCED };
    case "unavailable":
      return { type: VERDICTS.UNREPLAYABLE, reason: outcome.message };
    default: {
      const exhaustive: never = outcome;
      return exhaustive;
    }
  }
};

/** The base replay of `command`: where it runs, its arguments and its environment. */
export const baseReplay = (
  command: ReplayCommand & { file: string },
  base: string,
  env: NodeJS.ProcessEnv,
): { cwd: string; args: string[]; env: Record<string, string> } => {
  const inherited = Object.fromEntries(
    Object.entries(env).filter(
      (entry): entry is [string, string] =>
        entry[1] !== undefined && !STEP_ENV.some((key) => key === entry[0]),
    ),
  );
  return {
    cwd: path.join(base, command.dir),
    args: [
      "test",
      `./${command.file}`,
      ...(command.pattern === undefined ? [] : ["-t", command.pattern]),
    ],
    env: {
      ...inherited,
      // FOLIO_FAILURE lines, which tell a failed replay from a broken one,
      // are written under CI.
      CI: "true",
      [SEED_ENV]: String(command.seed),
      ...(command.path === null ? {} : { [PATH_ENV]: command.path }),
      ...(command.digest === undefined ? {} : { [CASE_ENV]: command.digest }),
      ...(command.factor === 1
        ? {}
        : { [NUM_RUNS_FACTOR_ENV]: String(command.factor) }),
    },
  };
};

type Runner = (replay: ReturnType<typeof baseReplay>) => ReplayOutcome;

const runReplay: Runner = ({ cwd, args, env }) => {
  console.log(`→ bun ${args.join(" ")} (in ${cwd})`);
  const result = spawnSync("bun", args, {
    cwd,
    env,
    encoding: "utf8",
    maxBuffer: 256 * 1024 * 1024,
    timeout: 30 * 60 * 1_000,
  });
  const output = `${result.stdout ?? ""}\n${result.stderr ?? ""}`;
  return replayOutcome({
    output,
    exitCode: result.status,
    ...(result.error && { error: result.error.message }),
  });
};

/** Each reported failure with what its base replay says. */
export const judgeFailures = ({
  failures,
  base,
  env,
  run = runReplay,
}: {
  failures: readonly ReportedFailure[];
  base: string;
  env: NodeJS.ProcessEnv;
  run?: Runner;
}): { failure: ReportedFailure; verdict: Verdict }[] =>
  failures.map((failure) => {
    const command = parseReplayCommand(failure.repro);
    if (command === undefined || command.file === undefined)
      return {
        failure,
        verdict: {
          type: VERDICTS.UNREPLAYABLE,
          reason: "its replay line names no test file to replay",
        },
      };
    const { file } = command;
    return {
      failure,
      verdict: verdictOf(run(baseReplay({ ...command, file }, base, env))),
    };
  });

/** The report lines and whether the change may land. */
export const summarize = (
  judged: readonly { failure: ReportedFailure; verdict: Verdict }[],
): { lines: string[]; blocks: boolean } => {
  const lines = judged.map(({ failure, verdict }) => {
    const name = `${failure.test} [${failure.fingerprint}]`;
    switch (verdict.type) {
      case VERDICTS.PRE_EXISTING:
        return `- ${name}: fails on the base too; filed as a finding. Replay: \`${failure.repro}\``;
      case VERDICTS.INTRODUCED:
        return `- ${name}: the base passes it, so this change brought it. Fix it or pin it with the fix. Replay: \`${failure.repro}\``;
      case VERDICTS.UNREPLAYABLE:
        return `- ${name}: could not be replayed on the base (${verdict.reason}), so nothing shows it was already there. Replay: \`${failure.repro}\``;
      default: {
        const exhaustive: never = verdict;
        return exhaustive;
      }
    }
  });
  return {
    lines,
    blocks: judged.some(
      ({ verdict }) => verdict.type !== VERDICTS.PRE_EXISTING,
    ),
  };
};

if (import.meta.main) {
  const args = process.argv.slice(2);
  const option = (name: string): string => {
    const index = args.indexOf(name);
    const value = index === -1 ? undefined : args.at(index + 1);
    return (
      value ??
      panic(
        "Usage: bun scripts/exploration-base-replay.ts --records <dir> --base <checkout>",
      )
    );
  };
  const records = option("--records");
  const base = option("--base");
  const failures = reportedFailures(records);
  if (failures.length === 0) {
    console.log("No generated property failure to replay on the base.");
    process.exit(0);
  }
  if (!existsSync(base)) panic(`The base checkout ${base} does not exist.`);
  const { lines, blocks } = summarize(
    judgeFailures({ failures, base, env: process.env }),
  );
  const report = [
    "## Generated property failures replayed on the base",
    "",
    ...lines,
  ].join("\n");
  console.log(report);
  const summaryFile = process.env["GITHUB_STEP_SUMMARY"];
  if (summaryFile !== undefined && summaryFile !== "")
    appendFileSync(summaryFile, `${report}\n`);
  if (blocks) {
    console.error(
      "::error::A generated property failure in a file this change does not own fails here but not on the base.",
    );
    process.exit(1);
  }
}
