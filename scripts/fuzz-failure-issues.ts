#!/usr/bin/env bun
/**
 * File fuzz failures as GitHub issues, one per failure class.
 *
 * Reads failure markers and minimized replay records. The class signature
 * groups a failure across fixtures and modes; fingerprints identify replay
 * cases in its seed table. Repeated runs are idempotent. An open issue of the
 * class always takes the rows. Recent closed classes reopen; older recurrences
 * link a new issue; one closed as a duplicate stands for the issue it names.
 * Known fingerprints remain tracked in their registry. A run that would open
 * more issues than the per-run cap opens none and exits non-zero.
 *
 * Usage:
 *   bun scripts/fuzz-failure-issues.ts [--log <file>]… [--records <dir>]…
 *     [--run-url <url>] [--sha <sha>] [--source <what ran>] [--dry-run]
 *
 * `--dry-run` prints the issues instead of calling `gh`.
 */

import { $ } from "bun";
import { TaggedError } from "better-result";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import type {
  FailureMarker,
  FailureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import { extractFailureMarkers, parseKnownFailures } from "./failure-fingerprints";
import { closedAsDuplicate, duplicateTarget, fileClasses, type Issue } from "./fuzz-issue-classes";

export const LABEL = {
  name: "fuzz-failure",
  color: "b60205",
  description: "A failure a fuzz run found, one issue per failure class",
};

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
    // An unshrunk report joins the shrunk finding it is the primary fingerprint of.
    const shrunk = [...findings.values()].find(
      ({ record }) => record.marker.primary === marker.fingerprint,
    );
    if (shrunk !== undefined) {
      if (!shrunk.seeds.includes(marker.seed)) shrunk.seeds.push(marker.seed);
      if (!shrunk.records.some(({ marker: existing }) => existing.seed === marker.seed)) {
        shrunk.records.push({
          version: 1,
          marker,
          replays: [marker.repro],
          error: marker.assertion,
        });
      }
      continue;
    }
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
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? "";
    if (arg === "--log") options.logs.push(value());
    else if (arg === "--records") options.records.push(value());
    else if (arg === "--run-url") options.runUrl = value() || null;
    else if (arg === "--sha") options.sha = value() || null;
    else if (arg === "--source") options.source = value() || options.source;
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
  return {
    number,
    title,
    body,
    state,
    closedAt,
    ...(typeof stateReason === "string" && stateReason !== "" ? { stateReason } : {}),
  };
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

/**
 * Open or update the issue of every finding (see the module comment);
 * returns where each went. `root` is the repository checkout.
 */
export const fileFindings = (
  findings: readonly Finding[],
  context: Context,
  root: string,
): Promise<Filed[]> => {
  const known = new Map(
    parseKnownFailures(
      JSON.parse(readFileSync(path.join(root, "test", "known-failure-fingerprints.json"), "utf8")),
    ).map((entry) => [entry.fingerprint, entry.issueOrPr]),
  );
  return fileClasses({
    findings,
    context,
    known,
    store: {
      list: async () => {
        const endpoint = `repos/{owner}/{repo}/issues?state=all&labels=${LABEL.name}&per_page=100`;
        const pages: unknown = JSON.parse(await $`gh api --paginate --slurp ${endpoint}`.text());
        const issues = parseIssuePages(pages);
        for (const issue of issues) {
          if (!closedAsDuplicate(issue)) continue;
          const comments = `repos/{owner}/{repo}/issues/${String(issue.number)}/comments?per_page=100`;
          const commentPages: unknown = JSON.parse(
            await $`gh api --paginate --slurp ${comments}`.text(),
          );
          issue.duplicateOf = duplicateTarget(parseCommentPages(commentPages));
        }
        return issues;
      },
      create: async (title, body) => {
        await $`gh label create ${LABEL.name} --color ${LABEL.color} --description ${LABEL.description} --force`.quiet();
        const url = (
          await $`gh issue create --title ${title} --label ${LABEL.name} --body-file ${writeBody(body)}`.text()
        ).trim();
        return parseIssueResponse(
          JSON.parse(await $`gh issue view ${url} --json number,title,state,body,closedAt`.text()),
        );
      },
      edit: async (number, body) => {
        await $`gh issue edit ${number} --body-file ${writeBody(body)}`.quiet();
      },
      reopen: async (number) => {
        await $`gh issue reopen ${number}`.quiet();
      },
    },
  });
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
  const findings = readFindings(options.logs, options.records);
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
    await fileClasses({
      findings,
      context,
      known: new Map(),
      maxNewIssues: Number.POSITIVE_INFINITY,
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
        edit: () => Promise.resolve(),
        reopen: () => Promise.resolve(),
      },
    });
    return;
  }
  const results = await fileFindings(findings, context, path.resolve(import.meta.dir, ".."));
  for (const { fingerprint, issue } of results)
    console.log(`${fingerprint}: ${issue ?? "not filed"}`);
};

if (import.meta.main) {
  await main();
}
