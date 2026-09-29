#!/usr/bin/env bun
/**
 * File fuzz failures as GitHub issues, one per fingerprint.
 *
 * Reads the `FOLIO_FAILURE {json}` lines of one or more logs and the failure
 * records a fuzz run wrote (`test/consumer-scenarios/support/failure-fingerprints.ts`,
 * which add the minimized flow and every replay line), groups them by
 * fingerprint, and for each one opens an issue labelled `fuzz-failure`, or
 * updates the issue already filed for it: an open one gets its occurrence
 * count, seeds and latest replay refreshed in place (no comment per run), a
 * closed one is reopened with a comment, because the failure came back. A
 * fingerprint `test/known-failure-fingerprints.json` lists is already
 * tracked where it says and gets no issue.
 *
 * Usage:
 *   bun scripts/fuzz-failure-issues.ts [--log <file>]… [--records <dir>]…
 *     [--run-url <url>] [--sha <sha>] [--source <what ran>] [--dry-run]
 *
 * `--dry-run` prints the issues instead of calling `gh`.
 */

import { $ } from "bun";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import type {
  FailureMarker,
  FailureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import { extractFailureMarkers, parseKnownFailures } from "./failure-fingerprints";

export const LABEL = {
  name: "fuzz-failure",
  color: "b60205",
  description: "A failure a fuzz run found, one issue per fingerprint",
};

/** New issues one run may open; updates to filed ones are not capped. */
const MAX_NEW_ISSUES = 15;
const MAX_TEXT = 6_000;
const MAX_SEEDS = 12;
const STATE_MARKER = "fuzz-failure-state";

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
export type Finding = { record: FailureRecord; seeds: number[] };

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
      findings.set(fingerprint, { record, seeds: [seed] });
      return;
    }
    if (!found.seeds.includes(seed)) found.seeds.push(seed);
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
      continue;
    }
    add({ version: 1, marker, replays: [marker.repro], error: marker.assertion }, true);
  }
  return [...findings.values()];
};

// ---------------------------------------------------------------------------
// Issue text
// ---------------------------------------------------------------------------

export type IssueState = {
  fingerprint: string;
  /** The fingerprint before shrinking, which an unshrunk report of the same failure carries. */
  primary?: string;
  count: number;
  firstSeen: string;
  lastSeen: string;
  seeds: number[];
};

export type Context = {
  runUrl: string | null;
  sha: string | null;
  source: string;
  date: string;
};

const truncate = (text: string): string =>
  text.length > MAX_TEXT
    ? `${text.slice(0, MAX_TEXT)}\n… (${String(text.length - MAX_TEXT)} more characters)`
    : text;

const fence = (text: string, lang = ""): string => {
  const ticks = text.includes("```") ? "````" : "```";
  return `${ticks}${lang}\n${truncate(text)}\n${ticks}`;
};

export const issueTitle = ({ record }: Finding): string => {
  const { fingerprint, test, assertion } = record.marker;
  const title = `Fuzz failure [${fingerprint}]: ${test}: ${assertion}`;
  return title.length <= 240 ? title : `${title.slice(0, 239)}…`;
};

/** The fingerprint an issue title files, or null. */
export const fingerprintOfTitle = (title: string): string | null =>
  /^Fuzz failure \[([0-9a-f]{16})\]/u.exec(title)?.[1] ?? null;

export const readState = (body: string): IssueState | null => {
  const match = new RegExp(`<!-- ${STATE_MARKER} (\\{.*?\\}) -->`, "u").exec(body);
  if (match === null) return null;
  try {
    const value: unknown = JSON.parse(match[1] as string);
    if (
      isRecord(value) &&
      typeof value["fingerprint"] === "string" &&
      typeof value["count"] === "number" &&
      typeof value["firstSeen"] === "string" &&
      typeof value["lastSeen"] === "string" &&
      Array.isArray(value["seeds"])
    ) {
      return value as IssueState;
    }
  } catch {
    // An edited state reads as none: the next update starts it afresh.
  }
  return null;
};

/** The state after this run saw `finding` (`previous` is the issue's, if filed). */
export const nextState = (
  finding: Finding,
  previous: IssueState | null,
  date: string,
): IssueState => ({
  fingerprint: previous?.fingerprint ?? finding.record.marker.fingerprint,
  primary: previous?.primary ?? finding.record.marker.primary ?? finding.record.marker.fingerprint,
  count: (previous?.count ?? 0) + 1,
  firstSeen: previous?.firstSeen ?? date,
  lastSeen: date,
  seeds: [...new Set([...finding.seeds, ...(previous?.seeds ?? [])])].slice(0, MAX_SEEDS),
});

export const issueBody = (finding: Finding, state: IssueState, context: Context): string => {
  const { record } = finding;
  const { marker } = record;
  const where = [
    context.runUrl === null ? context.source : `[${context.source}](${context.runUrl})`,
    context.sha === null ? null : `commit \`${context.sha.slice(0, 12)}\``,
  ]
    .filter((part) => part !== null)
    .join(", ");
  const [replay, ...others] = record.replays;
  const sections = [
    `A fuzz flow failed in ${where}.`,
    "",
    `**Test:** ${marker.test}`,
    `**Symptom:** \`${marker.assertion}\``,
    ...(marker.diff === undefined ? [] : [`**Differences:** \`${marker.diff}\``]),
    ...(marker.flow === undefined ? [] : [`**Minimized operations:** \`${marker.flow}\``]),
    `**Fingerprint:** \`${marker.fingerprint}\``,
    `**Seed:** \`${String(marker.seed)}\`${marker.path === null ? "" : `, **path:** \`${marker.path}\``}`,
    "",
    "### Replay",
    fence(replay ?? marker.repro, "sh"),
  ];
  if (others.length > 0) {
    sections.push("", "Also replays with:", fence(others.join("\n"), "sh"));
  }
  if (record.flow !== undefined) {
    const shrunk =
      record.shrink === undefined
        ? ""
        : ` (${String(record.shrink.steps)} steps, shrunk from ${String(record.shrink.from)} in ${String(record.shrink.attempts)} replays)`;
    sections.push(
      "",
      `### Minimized flow${shrunk}`,
      fence(JSON.stringify(record.flow, null, 2), "json"),
    );
  }
  if (record.error.trim() !== "") {
    sections.push("", "### Error", fence(record.error.trim()));
  }
  sections.push(
    "",
    "### Occurrences",
    `Seen in ${String(state.count)} run${state.count === 1 ? "" : "s"} since ${state.firstSeen}, last on ${state.lastSeen}. Seeds: ${state.seeds.map((seed) => `\`${String(seed)}\``).join(", ")}.`,
    "",
    `<!-- ${STATE_MARKER} ${JSON.stringify(state)} -->`,
  );
  return sections.join("\n");
};

// ---------------------------------------------------------------------------
// Filing
// ---------------------------------------------------------------------------

type Issue = { number: number; title: string; state: string; body: string | null };

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

/** Where a finding was filed: `#<number>`, or null when it was not. */
export type Filed = { fingerprint: string; issue: string | null };

/**
 * Open or update the issue of every finding (see the module comment);
 * returns where each went. `root` is the repository checkout.
 */
export const fileFindings = async (
  findings: readonly Finding[],
  context: Context,
  root: string,
): Promise<Filed[]> => {
  const known = new Map(
    parseKnownFailures(
      JSON.parse(readFileSync(path.join(root, "test", "known-failure-fingerprints.json"), "utf8")),
    ).map((entry) => [entry.fingerprint, entry]),
  );
  await $`gh label create ${LABEL.name} --color ${LABEL.color} --description ${LABEL.description} --force`
    .quiet()
    .nothrow();
  const endpoint = `repos/{owner}/{repo}/issues?state=all&labels=${LABEL.name}&per_page=100`;
  const pages = JSON.parse(await $`gh api --paginate --slurp ${endpoint}`.text()) as Issue[][];
  const filed = new Map<string, Issue>();
  const byPrimary = new Map<string, Issue>();
  const keep = (map: Map<string, Issue>, key: string, issue: Issue): void => {
    const current = map.get(key);
    if (current === undefined || (current.state !== "open" && issue.state === "open")) {
      map.set(key, issue);
    }
  };
  // Newest first, so an open issue wins over an older closed one.
  for (const issue of pages.flat().sort((a, b) => b.number - a.number)) {
    const fingerprint = fingerprintOfTitle(issue.title);
    if (fingerprint === null) continue;
    keep(filed, fingerprint, issue);
    keep(byPrimary, readState(issue.body ?? "")?.primary ?? fingerprint, issue);
  }
  // A shrunk failure and an unshrunk report of it share the primary fingerprint.
  const issueOf = ({ fingerprint, primary }: FailureMarker): Issue | undefined =>
    filed.get(fingerprint) ??
    (primary === undefined ? byPrimary.get(fingerprint) : filed.get(primary));

  const results: Filed[] = [];
  let opened = 0;
  for (const finding of findings) {
    const { fingerprint } = finding.record.marker;
    const title = issueTitle(finding);
    const { primary } = finding.record.marker;
    const entry =
      known.get(fingerprint) ?? (primary === undefined ? undefined : known.get(primary));
    const issue = issueOf(finding.record.marker);
    if (issue === undefined && entry !== undefined) {
      // Tracked where the registry says; a run every few hours adds nothing there.
      console.log(`known ${fingerprint}: ${entry.issueOrPr}`);
      results.push({ fingerprint, issue: entry.issueOrPr });
      continue;
    }
    if (issue === undefined) {
      if (opened >= MAX_NEW_ISSUES) {
        console.log(`not filed (cap of ${String(MAX_NEW_ISSUES)} new issues): ${title}`);
        results.push({ fingerprint, issue: null });
        continue;
      }
      const body = issueBody(finding, nextState(finding, null, context.date), context);
      const url = (
        await $`gh issue create --title ${title} --label ${LABEL.name} --body-file ${writeBody(body)}`.text()
      ).trim();
      opened += 1;
      console.log(`opened ${url}: ${title}`);
      const number = /\/issues\/(\d+)$/u.exec(url)?.[1];
      results.push({ fingerprint, issue: number === undefined ? null : `#${number}` });
      continue;
    }
    const number = String(issue.number);
    const body = issueBody(
      finding,
      nextState(finding, readState(issue.body ?? ""), context.date),
      context,
    );
    await $`gh issue edit ${number} --body-file ${writeBody(body)}`.quiet();
    if (issue.state !== "open") {
      await $`gh issue reopen ${number}`.quiet();
      const note = `Failed again in ${context.runUrl ?? context.source}, so reopening: the fix did not hold for this fingerprint.`;
      await $`gh issue comment ${number} --body-file ${writeBody(note)}`.quiet();
      console.log(`reopened #${number}: ${title}`);
    } else {
      console.log(`updated #${number}: ${title}`);
    }
    results.push({ fingerprint, issue: `#${number}` });
  }
  return results;
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
    date: new Date().toISOString().slice(0, 10),
  };
  if (findings.length === 0) {
    console.log("no fuzz failures to file");
    return;
  }
  if (options.dryRun) {
    for (const finding of findings) {
      console.log(
        `=== ${issueTitle(finding)}\n${issueBody(finding, nextState(finding, null, context.date), context)}\n`,
      );
    }
    return;
  }
  await fileFindings(findings, context, path.resolve(import.meta.dir, ".."));
};

if (import.meta.main) {
  await main();
}
