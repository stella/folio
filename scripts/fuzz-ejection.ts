#!/usr/bin/env bun
/**
 * Tell a merge group's fuzz failure that was already on its base from one the
 * group brought in, by replaying the group's failing seeds on the base.
 *
 *   replay  In a checkout of the base commit, rerun the consumer scenarios
 *           under the seed the group's run derived from its commit, and each
 *           failed property under its seed and counterexample path; write one
 *           log per replay. Runs the base's code, so it holds no write token.
 *   report  Compare the fingerprints the group's logs failed with to the ones
 *           the base replays failed with. Pre-existing ones get their issue
 *           (scripts/fuzz-failure-issues.ts); when every failure was
 *           pre-existing the pull request is labelled for another queue run.
 *           Prints one `FOLIO_QUEUE_EJECTION {json}` line either way.
 *
 * Usage:
 *   bun scripts/fuzz-ejection.ts replay --queue-logs <dir> --group-sha <sha>
 *     --base <checkout> --out <dir>
 *   bun scripts/fuzz-ejection.ts report --queue-logs <dir> --base-logs <dir>
 *     --group-sha <sha> --base-sha <sha> [--pr <n>] [--run-url <url>] [--dry-run]
 */

import { $ } from "bun";
import { spawnSync } from "node:child_process";
import { appendFileSync, existsSync, mkdirSync, readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { hash32 } from "../test/commit-seed";
import { extractFailureMarkers } from "./failure-fingerprints";
import { type Context, type Filed, fileFindings, readFindings } from "./fuzz-failure-issues";

export const LABEL = {
  name: "queue-ejection-pre-existing",
  color: "fbca04",
  description: "Ejected from the merge queue by a failure its base already had",
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

const run = (
  command: string,
  args: string[],
  { cwd, env, log }: { cwd: string; env: Record<string, string>; log: string },
): void => {
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
};

const replay = (options: Options): void => {
  if (!SHA.test(options.groupSha)) throw new Error("--group-sha must be a commit sha");
  mkdirSync(options.out, { recursive: true });
  const logs = logsIn(options.queueLogs).map((file) => readFileSync(file, "utf8"));
  const markers = logs.flatMap((log) => extractFailureMarkers(log));
  if (markers.some((marker) => marker.repro.includes("consumer-scenarios.ts"))) {
    // The group's run seeded its fuzz flows from the group commit; the base
    // replays the same flows under that seed, with the same defaults.
    run("bun", ["scripts/consumer-scenarios.ts"], {
      cwd: options.base,
      env: { FOLIO_SCENARIO_SEED: String(consumerSeed(options.groupSha)) },
      log: path.join(options.out, "base-consumer-scenarios.log"),
    });
  }
  const seen = new Set<string>();
  logs.flatMap(propertyReplays).forEach((property, index) => {
    const key = `${property.file}\0${property.title}\0${String(property.seed)}`;
    if (seen.has(key)) return;
    seen.add(key);
    const { dir, file } = packageOf(property.file);
    run("bun", ["test", file, "-t", titlePattern(property.title)], {
      cwd: path.join(options.base, dir),
      env: {
        CI: "true",
        PROPERTY_TEST_SEED: String(property.seed),
        PROPERTY_TEST_NUM_RUNS_FACTOR: String(property.factor),
        ...(property.path === null ? {} : { PROPERTY_TEST_PATH: property.path }),
      },
      log: path.join(options.out, `base-property-${String(index)}.log`),
    });
  });
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

/** Which of the group's failures the base replays failed with too. */
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

const report = async (options: Options): Promise<void> => {
  if (!SHA.test(options.groupSha) || !SHA.test(options.baseSha)) {
    throw new Error("--group-sha and --base-sha must be commit shas");
  }
  const findings = readFindings(logsIn(options.queueLogs), []);
  const base = new Set(
    logsIn(options.baseLogs).flatMap((file) =>
      extractFailureMarkers(readFileSync(file, "utf8")).map((marker) => marker.fingerprint),
    ),
  );
  const judged = judge(
    findings.map(({ record }) => record.marker),
    base,
  );
  const ejection: Ejection = {
    pr: options.pr,
    group: options.groupSha,
    base: options.baseSha,
    ...judged,
  };
  const preExisting = findings.filter(({ record }) => base.has(record.marker.fingerprint));
  if (!options.dryRun && preExisting.length > 0) {
    const context: Context = {
      runUrl: options.runUrl,
      sha: options.baseSha,
      source: "a merge queue run, and again on its base",
      date: new Date().toISOString().slice(0, 10),
    };
    const filed: Filed[] = await fileFindings(
      preExisting,
      context,
      path.resolve(import.meta.dir, ".."),
    );
    for (const entry of ejection.fingerprints) {
      entry.issue =
        filed.find(({ fingerprint }) => fingerprint === entry.fingerprint)?.issue ?? null;
    }
  }
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
            `- \`${entry.fingerprint}\` ${entry.test}: ${entry.reproducedOnBase ? "fails on the base too" : "passes on the base"}${entry.issue === null ? "" : ` (${entry.issue})`}`,
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
