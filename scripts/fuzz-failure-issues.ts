#!/usr/bin/env bun
/**
 * File one issue per fingerprint. Each recurrence adds a bounded run comment;
 * closed issues reopen. Issue bodies never accumulate replay history.
 */

import { $ } from "bun";
import { panic, TaggedError } from "better-result";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import type {
  FailureMarker,
  FailureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import { extractFailureMarkers } from "./failure-fingerprints";
import { type ClosedIssueDisposition, type Issue, type IssueStore } from "./fuzz-issue-classes";

export const LABEL = {
  name: "fuzz-failure",
  color: "b60205",
  description: "A failure a fuzz run found, one issue per fingerprint",
};

export const TITLE_PREFIX = "Fuzz failure";
export const TITLE_FORMAT = `${TITLE_PREFIX} [<fingerprint>]: <test>`;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** A failure record read from disk, or null when it is not one. */
export const parseFailureRecord = (value: unknown): FailureRecord | null => {
  if (!isRecord(value) || value["version"] !== 1 || !isRecord(value["marker"])) return null;
  let marker: FailureMarker | undefined;
  try {
    [marker] = extractFailureMarkers(`FOLIO_FAILURE ${JSON.stringify(value["marker"])}`);
  } catch {
    return null;
  }
  const replays = value["replays"];
  if (
    marker === undefined ||
    !Array.isArray(replays) ||
    replays.length === 0 ||
    !replays.every((replay) => typeof replay === "string" && !replay.includes("\n")) ||
    typeof value["error"] !== "string"
  ) {
    return null;
  }
  const shrink = value["shrink"];
  return {
    version: 1,
    marker,
    replays: replays as string[],
    error: value["error"],
    ...(value["flow"] === undefined ? {} : { flow: value["flow"] }),
    ...(isRecord(shrink) &&
    typeof shrink["steps"] === "number" &&
    typeof shrink["from"] === "number" &&
    typeof shrink["attempts"] === "number"
      ? { shrink: { steps: shrink["steps"], from: shrink["from"], attempts: shrink["attempts"] } }
      : {}),
  };
};

/** One fingerprint's failures from a run: the most useful record and every seed. */
export type Finding = { record: FailureRecord; seeds: number[]; records: FailureRecord[] };

const stepsOf = (record: FailureRecord): number => record.shrink?.steps ?? Number.POSITIVE_INFINITY;

/**
 * Group records and bare markers by fingerprint. A record beats a bare
 * marker, and a shorter minimized flow beats a longer one.
 */
export const collectFindings = (
  records: readonly FailureRecord[],
  markers: readonly FailureMarker[],
): Finding[] => {
  const findings = new Map<string, Finding>();
  const add = (record: FailureRecord, bare: boolean): void => {
    const { fingerprint, seed } = record.marker;
    const found = findings.get(fingerprint);
    if (found === undefined) {
      findings.set(fingerprint, { record, seeds: [seed], records: [record] });
      return;
    }
    if (!found.seeds.includes(seed)) found.seeds.push(seed);
    const existingIndex = found.records.findIndex(
      ({ marker }) => marker.seed === seed && marker.path === record.marker.path,
    );
    if (existingIndex < 0) found.records.push(record);
    else {
      const existing = found.records.at(existingIndex);
      if (existing !== undefined && !bare && stepsOf(record) < stepsOf(existing)) {
        found.records[existingIndex] = record;
      }
    }
    if (!bare && stepsOf(record) < stepsOf(found.record)) found.record = record;
  };
  for (const record of records) add(record, false);
  for (const marker of markers) {
    add({ version: 1, marker, replays: [marker.repro], error: marker.assertion }, true);
  }
  return [...findings.values()];
};

export type Context = {
  runUrl: string | null;
  sha: string | null;
  source: string;
  date: string;
};

const parseArgs = (argv: readonly string[]) => {
  const options = {
    logs: [] as string[],
    records: [] as string[],
    runUrl: null as string | null,
    sha: null as string | null,
    source: "a fuzz run",
    dryRun: false,
    requireFailures: false,
    artifactRoot: null as string | null,
    logPattern: "**/*.log",
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? "";
    if (arg === "--log") options.logs.push(value());
    else if (arg === "--records") options.records.push(value());
    else if (arg === "--run-url") options.runUrl = value() || null;
    else if (arg === "--sha") options.sha = value() || null;
    else if (arg === "--source") options.source = value() || options.source;
    else if (arg === "--artifact-root") options.artifactRoot = value();
    else if (arg === "--log-pattern") options.logPattern = value();
    else if (arg === "--require-failures") options.requireFailures = true;
    else if (arg === "--dry-run") options.dryRun = true;
    else throw new Error(`Unknown argument ${String(arg)}`);
  }
  return options;
};

const readRecords = (dir: string): FailureRecord[] => {
  if (!existsSync(dir)) return [];
  return readdirSync(dir, { recursive: true })
    .map(String)
    .filter((file) => file.endsWith(".json"))
    .flatMap((file) => {
      try {
        const record = parseFailureRecord(JSON.parse(readFileSync(path.join(dir, file), "utf8")));
        return record === null ? [] : [record];
      } catch {
        return [];
      }
    });
};

const writeBody = (body: string): string => {
  const file = path.join(tmpdir(), `fuzz-failure-${String(process.pid)}.md`);
  writeFileSync(file, body);
  return file;
};

class IssueResponseError extends TaggedError("IssueResponseError")<{ message: string }> {}

/** GitHub's close reasons (REST `state_reason`, GraphQL `stateReason`) a closed issue can carry. */
const CLOSE_REASONS = {
  completed: { type: "fixed" },
  not_planned: { type: "superseded" },
  duplicate: { type: "duplicate" },
} as const satisfies Record<string, ClosedIssueDisposition>;

const isCloseReason = (reason: string): reason is keyof typeof CLOSE_REASONS =>
  Object.hasOwn(CLOSE_REASONS, reason);

/** A closed issue's disposition; a close reason this reporter does not know fails. */
const closedIssueDisposition = (stateReason: unknown): ClosedIssueDisposition => {
  // Issues closed before GitHub recorded close reasons carry none: they were closed as done.
  if (stateReason === null || stateReason === undefined || stateReason === "") {
    return CLOSE_REASONS.completed;
  }
  const reason = typeof stateReason === "string" ? stateReason.toLowerCase() : null;
  if (reason === null || !isCloseReason(reason)) {
    throw new IssueResponseError({
      message: `Unknown close reason ${JSON.stringify(stateReason)} on a closed issue`,
    });
  }
  return CLOSE_REASONS[reason];
};

/** Reject malformed API responses before they can affect matching or writes. */
export const parseIssueResponse = (value: unknown): Issue => {
  if (!isRecord(value)) throw new IssueResponseError({ message: "Invalid GitHub issue response" });
  const number = value["number"];
  const title = value["title"];
  const body = value["body"];
  const state = typeof value["state"] === "string" ? value["state"].toLowerCase() : null;
  const closedAt = value["closed_at"] ?? value["closedAt"] ?? null;
  const stateReason = value["state_reason"] ?? value["stateReason"] ?? null;
  if (
    typeof number !== "number" ||
    !Number.isSafeInteger(number) ||
    number <= 0 ||
    typeof title !== "string" ||
    (body !== null && typeof body !== "string") ||
    (state !== "open" && state !== "closed") ||
    (closedAt !== null && typeof closedAt !== "string")
  )
    throw new IssueResponseError({ message: "Invalid GitHub issue fields" });
  if (state === "open") return { number, title, body, state, closedAt: null };
  return { number, title, body, state, closedAt, disposition: closedIssueDisposition(stateReason) };
};

/** The bodies of an issue's comments, from paginated GitHub responses. */
export const parseCommentPages = (pages: unknown): string[] => {
  if (!Array.isArray(pages) || !pages.every(Array.isArray)) {
    throw new IssueResponseError({ message: "Invalid GitHub comment pages" });
  }
  return pages.flat().map((comment: unknown) => {
    if (!isRecord(comment) || typeof comment["body"] !== "string") {
      throw new IssueResponseError({ message: "Invalid GitHub comment fields" });
    }
    return comment["body"];
  });
};

/** Validate paginated GitHub responses before selecting an issue to update. */
export const parseIssuePages = (pages: unknown): Issue[] => {
  if (!Array.isArray(pages) || !pages.every(Array.isArray)) {
    throw new IssueResponseError({ message: "Invalid GitHub issue pages" });
  }
  return pages.flat().map(parseIssueResponse);
};

/** Where a finding was filed: `#<number>`, or null when it was not. */
export type Filed = { fingerprint: string; issue: string | null };

export type FingerprintIssueStore = Pick<IssueStore, "list" | "create"> & {
  reopen: (number: number) => Promise<void>;
  comment: (number: number, body: string) => Promise<void>;
};

class MissingFailuresError extends TaggedError("MissingFailuresError")<{ message: string }> {}

export const requireFindings = (findings: readonly Finding[]): void => {
  if (findings.length === 0) {
    throw new MissingFailuresError({
      message:
        "Failed fuzz/flow run contained 0 FOLIO_FAILURE fingerprints; check the failure artifact and marker producer",
    });
  }
};

// Bound each field so even a large assertion or replay stays below GitHub's body limit.
const FIELD_LIMIT = 4_000;
const field = (value: string, limit = FIELD_LIMIT): string => value.slice(0, limit);
const code = (value: string, limit = FIELD_LIMIT): string => {
  const bounded = field(value, limit);
  const fence = "`".repeat(
    Math.max(3, ...[...bounded.matchAll(/`+/gu)].map(([run]) => run.length + 1)),
  );
  return `${fence}\n${bounded}\n${fence}`;
};
const bodyOf = ({ record: { marker } }: Finding): string =>
  [
    `Fingerprint: ${marker.fingerprint}`,
    `Test: ${field(marker.test)}`,
    `Assertion:\n${code(marker.assertion)}`,
  ].join("\n\n");

export const MAX_RUN_SEEDS = 10;
const REPLAY_FIELD_LIMIT = 1_000;

type RunEvidenceOptions = { finding: Finding; context: Context };
const runEvidence = ({ finding, context }: RunEvidenceOptions): string => {
  const seeds = finding.seeds.slice(0, MAX_RUN_SEEDS);
  const omitted = finding.seeds.length - seeds.length;
  const replays = seeds.map((seed) => {
    const record = finding.records.find(({ marker: item }) => item.seed === seed);
    if (record === undefined) panic(`Missing replay record for seed ${seed}`);
    return `Seed: ${seed}\nRepro command:\n${code(record.marker.repro, REPLAY_FIELD_LIMIT)}`;
  });
  return [
    `Run: ${field(context.runUrl ?? "unavailable")}`,
    `Seeds: ${seeds.join(", ")}`,
    ...(omitted > 0 ? [`${omitted} more seeds; see run artifacts.`] : []),
    bodyOf(finding),
    ...replays,
  ].join("\n\n");
};

const fingerprintOf = (issue: Issue): string | undefined =>
  issue.body?.match(/^Fingerprint: ([0-9a-f]{16})$/mu)?.at(1) ??
  new RegExp(`^${TITLE_PREFIX} \\[([0-9a-f]{16})\\]`, "u").exec(issue.title)?.at(1);

type FileFindingsOptions = {
  findings: readonly Finding[];
  context: Context;
  store?: FingerprintIssueStore;
};
export const fileFindings = async ({
  findings,
  context,
  store,
}: FileFindingsOptions): Promise<Filed[]> => {
  const client =
    store ??
    ({
      list: async () =>
        parseIssuePages(
          JSON.parse(
            await $`gh api --paginate --slurp ${`repos/{owner}/{repo}/issues?state=all&labels=${LABEL.name}&per_page=100`}`.text(),
          ),
        ),
      create: async (title, body) => {
        await $`gh label create ${LABEL.name} --color ${LABEL.color} --description ${LABEL.description} --force`.quiet();
        const url = (
          await $`gh issue create --title ${title} --label ${LABEL.name} --body-file ${writeBody(body)}`.text()
        ).trim();
        return parseIssueResponse(
          JSON.parse(
            await $`gh issue view ${url} --json number,title,state,body,closedAt,stateReason`.text(),
          ),
        );
      },
      reopen: async (number) => {
        await $`gh issue reopen ${number}`.quiet();
      },
      comment: async (number, body) => {
        await $`gh issue comment ${number} --body-file ${writeBody(body)}`.quiet();
      },
    } satisfies FingerprintIssueStore);
  const issues = [...(await client.list())];
  const filed: Filed[] = [];
  // A caller may hand us repeated records; identity always comes from the fingerprint.
  const unique = collectFindings(
    findings.flatMap(({ records }) => records),
    findings.map(({ record }) => record.marker),
  );
  for (const finding of unique) {
    const { fingerprint, test } = finding.record.marker;
    let issue =
      issues.find(
        (candidate) => fingerprintOf(candidate) === fingerprint && candidate.state === "open",
      ) ?? issues.find((candidate) => fingerprintOf(candidate) === fingerprint);
    const evidence = runEvidence({ finding, context });
    if (issue === undefined) {
      issue = await client.create(
        `${TITLE_PREFIX} [${fingerprint}]: ${field(test).slice(0, 200)}`,
        evidence,
      );
      issues.push(issue);
    } else {
      if (issue.state === "closed") await client.reopen(issue.number);
      await client.comment(issue.number, evidence);
    }
    filed.push({ fingerprint, issue: `#${issue.number}` });
  }
  return filed;
};

/** Discover selected logs without merging matrix artifacts with identical filenames. */
export const artifactInputs = (root: string, pattern = "**/*.log") => {
  const logs = [...new Bun.Glob(pattern).scanSync({ cwd: root, onlyFiles: true })].map((file) =>
    path.join(root, file),
  );
  return { logs, records: [...new Set(logs.map((file) => path.dirname(file)))] };
};

/** The findings in `logs` (their FOLIO_FAILURE lines) and the record directories. */
export const readFindings = (logs: readonly string[], records: readonly string[]): Finding[] =>
  collectFindings(
    records.flatMap(readRecords),
    logs
      .filter((log) => existsSync(log))
      .flatMap((log) => extractFailureMarkers(readFileSync(log, "utf8"))),
  );

const main = async (): Promise<void> => {
  const options = parseArgs(process.argv.slice(2));
  const artifacts =
    options.artifactRoot === null
      ? { logs: [], records: [] }
      : artifactInputs(options.artifactRoot, options.logPattern);
  const findings = readFindings(
    [...options.logs, ...artifacts.logs],
    [...options.records, ...artifacts.records],
  );
  if (options.requireFailures) requireFindings(findings);
  const context: Context = {
    runUrl: options.runUrl,
    sha: options.sha,
    source: options.source,
    date: new Date().toISOString(),
  };
  if (findings.length === 0) {
    console.log("no fuzz failures to file");
    return;
  }
  if (options.dryRun) {
    let number = 0;
    await fileFindings({
      findings,
      context,
      store: {
        list: () => Promise.resolve([]),
        create: (title, body) => {
          console.log(`=== ${title}\n${body}\n`);
          return Promise.resolve({
            number: ++number,
            title,
            body,
            state: "open",
            closedAt: null,
          } satisfies Issue);
        },
        comment: () => Promise.resolve(),
        reopen: () => Promise.resolve(),
      },
    });
    return;
  }
  const results = await fileFindings({
    findings,
    context,
  });
  for (const { fingerprint, issue } of results)
    console.log(`${fingerprint}: ${issue ?? "not filed"}`);
};

if (import.meta.main) {
  await main();
}
