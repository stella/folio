import { Result, TaggedError } from "better-result";
import {
  classOfTitle,
  failureClass,
  type FailureClass,
  upgradeClassKey,
} from "./fuzz-failure-class";
import type { Context, Filed, Finding } from "./fuzz-failure-issues";

export type Issue = {
  number: number;
  title: string;
  state: "open" | "closed";
  body: string | null;
  closedAt: string | null;
  /** Why a closed issue was closed, as GitHub reports it. */
  stateReason?: string | null;
  /** The issue a "Duplicate of #N" comment names, for one closed as a duplicate. */
  duplicateOf?: number | null;
};

export type IssueStore = {
  list: () => Promise<Issue[]>;
  create: (title: string, body: string) => Promise<Issue>;
  edit: (number: number, body: string) => Promise<void>;
  reopen: (number: number) => Promise<void>;
};

/** A run found more new classes than the reporter may open issues for. */
export class NewIssueCapError extends TaggedError("NewIssueCapError")<{ message: string }> {}

type SeedRow = {
  fingerprint: string;
  fixture?: string;
  flow: string;
  mode: string;
  seed: number;
  path: string | null;
  repro: string;
  replays: string[];
  evidence: string;
  firstSeenRun: string;
};
type ClassState = { key: string; runs: string[]; seeds: SeedRow[] };
const STATE_MARKER = "fuzz-class-state";
const REOPEN_WINDOW_MS = 14 * 24 * 60 * 60 * 1_000;
export const MAX_NEW_ISSUES = 5;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isSeedRow = (value: unknown): value is SeedRow =>
  isRecord(value) &&
  typeof value["fingerprint"] === "string" &&
  /^[0-9a-f]{16}$/u.test(value["fingerprint"]) &&
  (value["fixture"] === undefined || typeof value["fixture"] === "string") &&
  typeof value["flow"] === "string" &&
  typeof value["mode"] === "string" &&
  typeof value["seed"] === "number" &&
  Number.isSafeInteger(value["seed"]) &&
  (value["path"] === null || typeof value["path"] === "string") &&
  typeof value["repro"] === "string" &&
  !value["repro"].includes("\n") &&
  Array.isArray(value["replays"]) &&
  value["replays"].every(
    (replay: unknown) => typeof replay === "string" && !replay.includes("\n"),
  ) &&
  typeof value["evidence"] === "string" &&
  typeof value["firstSeenRun"] === "string";

/** A stored key in today's form, or undefined when it is not a key. */
const currentKey = (json: string): string | undefined => {
  const parsed = Result.try((): unknown => JSON.parse(json));
  if (parsed.isErr()) return undefined;
  const parts: unknown = parsed.value;
  if (!Array.isArray(parts) || !parts.every((part: unknown) => typeof part === "string")) {
    return undefined;
  }
  return upgradeClassKey(parts);
};

const readClassState = (body: string | null, key: string): ClassState | null => {
  const json = body?.match(/<!-- fuzz-class-state (\{[^\n]*\}) -->/u)?.at(1);
  if (json === undefined) return null;
  const parsed = Result.try((): unknown => JSON.parse(json));
  if (parsed.isErr()) return null;
  const value = parsed.value;
  if (
    !isRecord(value) ||
    typeof value["key"] !== "string" ||
    currentKey(value["key"]) !== key ||
    !Array.isArray(value["runs"]) ||
    !value["runs"].every((run: unknown) => typeof run === "string") ||
    !Array.isArray(value["seeds"]) ||
    !value["seeds"].every(isSeedRow)
  )
    return null;
  return { key, runs: [...new Set(value["runs"])], seeds: value["seeds"] };
};

const markedKey = (issue: Issue): string | undefined => {
  const marker = issue.body?.match(/<!-- fuzz-class: (.*?) -->/u)?.at(1);
  return marker === undefined ? undefined : currentKey(marker);
};

const legacyIdentity = (issue: Issue): string[] => {
  const marker = issue.body?.match(/<!-- fuzz-failure-state (\{[^\n]*\}) -->/u)?.at(1);
  if (marker === undefined) return [];
  const parsed = Result.try((): unknown => JSON.parse(marker));
  if (parsed.isErr() || !isRecord(parsed.value)) return [];
  return [parsed.value["fingerprint"], parsed.value["primary"]].filter(
    (value): value is string => typeof value === "string" && /^[0-9a-f]{16}$/u.test(value),
  );
};
const legacyKey = (issue: Issue): string | undefined => classOfTitle(issue.title)?.key;

const lowestNumbered = (issues: readonly Issue[]): Issue | undefined =>
  issues.toSorted((a, b) => a.number - b.number).at(0);

const latestClosed = (issues: readonly Issue[]): Issue | undefined =>
  issues
    .toSorted(
      (a, b) =>
        (Date.parse(b.closedAt ?? "") || 0) - (Date.parse(a.closedAt ?? "") || 0) ||
        b.number - a.number,
    )
    .at(0);

/** Closed as not planned (or as a duplicate): the closures a duplicate comment redirects. */
export const closedAsDuplicate = ({ state, stateReason }: Issue): boolean =>
  state === "closed" && /^(?:not_planned|duplicate)$/iu.test(stateReason ?? "");

/** The issue the last "Duplicate of #N" comment names. */
export const duplicateTarget = (comments: readonly string[]): number | null => {
  const target = comments
    .flatMap((comment) => [...comment.matchAll(/^\s*Duplicate of #(\d+)\b/gimu)])
    .at(-1)
    ?.at(1);
  return target === undefined ? null : Number(target);
};

/**
 * Follow duplicate closures to the issue that stands for them: an issue of
 * the list, or only its number when the list does not have it.
 */
const duplicateEnd = (issue: Issue, byNumber: ReadonlyMap<number, Issue>): Issue | number => {
  const seen = new Set<number>();
  let current = issue;
  while (closedAsDuplicate(current) && typeof current.duplicateOf === "number") {
    seen.add(current.number);
    if (seen.has(current.duplicateOf)) break;
    const next = byNumber.get(current.duplicateOf);
    if (next === undefined) return current.duplicateOf;
    current = next;
  }
  return current;
};

type ClassGroup = {
  signature: FailureClass;
  rows: SeedRow[];
  /** The findings of the class. */
  fingerprints: Set<string>;
  /** Every fingerprint an issue that predates class markers may know a row by. */
  identities: Set<string>;
};

/** The issue a class reports to, or the number of one it may only point at. */
const classIssue = (group: ClassGroup, issues: readonly Issue[]): Issue | number | undefined => {
  const { key } = group.signature;
  // In order: the class marker; then, on issues without one, an exact
  // fingerprint, then the failure message of the title.
  const matching = issues.filter((issue) => {
    const marked = markedKey(issue);
    if (marked !== undefined) return marked === key;
    return (
      legacyIdentity(issue).some((identity) => group.identities.has(identity)) ||
      legacyKey(issue) === key
    );
  });
  const open = lowestNumbered(matching.filter(({ state }) => state === "open"));
  if (open !== undefined) return open;
  const byNumber = new Map(issues.map((issue) => [issue.number, issue]));
  const ends = matching.map((issue) => duplicateEnd(issue, byNumber));
  // A duplicate's target that reports another class is pointed at, never rewritten.
  const foreign = (end: Issue | number): boolean => {
    if (typeof end === "number") return true;
    const marked = markedKey(end);
    return marked !== undefined && marked !== key;
  };
  const own = [...new Set(ends.filter((end): end is Issue => !foreign(end)))];
  const openEnd = lowestNumbered(own.filter(({ state }) => state === "open"));
  if (openEnd !== undefined) return openEnd;
  const elsewhere = ends
    .filter(foreign)
    .map((end) => (typeof end === "number" ? end : end.number))
    .toSorted((a, b) => a - b)
    .at(0);
  return elsewhere ?? latestClosed(own);
};

const rowIdentity = ({ fingerprint, seed, path }: SeedRow): string =>
  JSON.stringify([fingerprint, seed, path]);
const tableCell = (value: string): string =>
  value.replaceAll("|", "\\|").replaceAll("\n", " ").replaceAll("`", "\\`");

type ClassBodyOptions = {
  signature: FailureClass;
  state: ClassState;
  note: string | null;
  legacyReport: string | null;
};
const classBody = ({ signature, state, note, legacyReport }: ClassBodyOptions): string => {
  const rows = state.seeds.map(
    (row) =>
      `| ${row.fingerprint} | ${tableCell(row.fixture ?? "")} | ${tableCell(row.flow)} | ${tableCell(row.mode)} | ${row.seed} | ${tableCell(row.path ?? "")} | ${tableCell(row.repro)} | ${tableCell(row.firstSeenRun)} |`,
  );
  const json = JSON.stringify(state).replaceAll("<", "\\u003c").replaceAll(">", "\\u003e");
  return [
    `Class signature: \`${signature.key}\``,
    `Failure: ${signature.message}`,
    `Seen in ${state.runs.length} distinct runs; ${state.seeds.length} replay rows.`,
    ...(note === null ? [] : [note]),
    "",
    "| Fingerprint | Fixture | Flow | Mode | Seed | Path | One-line repro | First seen run |",
    "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ...rows,
    "",
    ...state.seeds.flatMap((row) => [
      ...row.replays
        .slice(1)
        .map(
          (replay) =>
            `Additional replay for ${row.fingerprint}, seed ${row.seed}: ${tableCell(replay)}`,
        ),
      `<details><summary>Evidence for ${row.fingerprint}, seed ${row.seed}</summary>`,
      "",
      row.evidence
        .split("\n")
        .map((line) => `    ${line}`)
        .join("\n"),
      "",
      "</details>",
    ]),
    "",
    `<!-- fuzz-class: ${signature.key.replaceAll("<", "\\u003c").replaceAll(">", "\\u003e")} -->`,
    `<!-- ${STATE_MARKER} ${json} -->`,
    ...(legacyReport === null
      ? []
      : ["", "<details><summary>Legacy report</summary>", "", legacyReport, "", "</details>"]),
  ].join("\n");
};

type FileClassesOptions = {
  findings: readonly Finding[];
  context: Context;
  store: IssueStore;
  known: ReadonlyMap<string, string>;
  maxNewIssues?: number;
};
export const fileClasses = async ({
  findings,
  context,
  store,
  known,
  maxNewIssues = MAX_NEW_ISSUES,
}: FileClassesOptions): Promise<Filed[]> => {
  const run = context.runUrl ?? `${context.source}:${context.sha ?? context.date}`;
  const filed = new Map<string, Set<string>>();
  const groups = new Map<string, ClassGroup>();
  const hasUntracked = findings.some(
    ({ record: { marker } }) =>
      !known.has(marker.fingerprint) && !known.has(marker.primary ?? marker.fingerprint),
  );
  const issues = hasUntracked ? await store.list() : [];
  for (const finding of findings) {
    const fingerprint = finding.record.marker.fingerprint;
    const tracked =
      known.get(fingerprint) ?? known.get(finding.record.marker.primary ?? fingerprint);
    if (tracked !== undefined) {
      filed.set(fingerprint, new Set([tracked]));
      continue;
    }
    for (const record of finding.records) {
      const signature = failureClass(record.marker.test, record.marker.assertion);
      let group = groups.get(signature.key);
      if (group === undefined) {
        group = { signature, rows: [], fingerprints: new Set(), identities: new Set() };
        groups.set(signature.key, group);
      }
      group.fingerprints.add(fingerprint);
      group.identities.add(record.marker.fingerprint);
      if (record.marker.primary !== undefined) group.identities.add(record.marker.primary);
      const replays = record.replays.length > 0 ? record.replays : [record.marker.repro];
      group.rows.push({
        fingerprint: record.marker.fingerprint,
        fixture: signature.fixture,
        flow: record.marker.flow ?? signature.flow,
        mode: signature.mode,
        seed: record.marker.seed,
        path: record.marker.path,
        repro: replays.at(0) ?? record.marker.repro,
        replays: [...replays],
        evidence: [
          record.error,
          record.marker.diff ?? "",
          record.flow === undefined ? "" : JSON.stringify(record.flow).slice(0, 600),
        ].join("\n"),
        firstSeenRun: run,
      });
    }
  }
  const file = (fingerprints: ReadonlySet<string>, issue: number): void => {
    for (const fingerprint of fingerprints) {
      const references = filed.get(fingerprint) ?? new Set<string>();
      references.add(`#${issue}`);
      filed.set(fingerprint, references);
    }
  };
  const pending: { title: string; body: string; fingerprints: Set<string> }[] = [];
  for (const group of groups.values()) {
    const { signature, rows, fingerprints } = group;
    const previous = classIssue(group, issues);
    if (typeof previous === "number") {
      file(fingerprints, previous);
      continue;
    }
    const recent =
      previous?.state === "closed" &&
      previous.closedAt !== null &&
      Date.parse(context.date) - Date.parse(previous.closedAt) >= 0 &&
      Date.parse(context.date) - Date.parse(previous.closedAt) < REOPEN_WINDOW_MS;
    const reuse = previous !== undefined && (previous.state === "open" || recent);
    const previousState = reuse ? readClassState(previous.body, signature.key) : null;
    let legacyReport: string | null = null;
    if (reuse) {
      if (previousState === null) legacyReport = previous.body;
      else
        legacyReport =
          previous.body
            ?.match(/<details><summary>Legacy report<\/summary>\n\n([\s\S]*?)\n\n<\/details>/u)
            ?.at(1) ?? null;
    }
    const state = previousState ?? { key: signature.key, runs: [], seeds: [] };
    let changed = false;
    if (!state.runs.includes(run)) {
      state.runs.push(run);
      changed = true;
    }
    const identities = new Map(state.seeds.map((row) => [rowIdentity(row), row]));
    for (const row of rows) {
      const identity = rowIdentity(row);
      const existing = identities.get(identity);
      if (existing !== undefined) {
        for (const replay of row.replays) {
          if (existing.replays.includes(replay)) continue;
          existing.replays.push(replay);
          changed = true;
        }
        continue;
      }
      identities.set(identity, row);
      state.seeds.push(row);
      changed = true;
    }
    const oldNote = previous?.body
      ?.split("\n")
      .find((line) => line.startsWith("Regressed:") || line.startsWith("Previous occurrence:"));
    let note: string | null = null;
    if (recent) note = `Regressed: this class recurred in ${run}.`;
    else if (reuse) note = oldNote ?? null;
    else if (previous !== undefined) note = `Previous occurrence: #${previous.number}.`;
    const body = classBody({ signature, state, note, legacyReport });
    if (!reuse) {
      const { title } = signature;
      pending.push({
        title: title.length > 240 ? `${title.slice(0, 239)}…` : title,
        body,
        fingerprints,
      });
      continue;
    }
    if (recent) {
      await store.reopen(previous.number);
      previous.state = "open";
      previous.closedAt = null;
    }
    if (changed || recent) {
      previous.body = body;
      await store.edit(previous.number, body);
    }
    file(fingerprints, previous.number);
  }
  // Existing issues are updated above whatever happens here. A run that would
  // open more issues than the cap opens none and fails instead.
  if (pending.length > maxNewIssues) {
    throw new NewIssueCapError({
      message: [
        `Refusing to open ${pending.length} fuzz failure issues in one run (limit ${maxNewIssues}); none were opened.`,
        ...pending.map(({ title }) => `  ${title}`),
      ].join("\n"),
    });
  }
  for (const { title, body, fingerprints } of pending) {
    const issue = await store.create(title, body);
    file(fingerprints, issue.number);
  }
  return [...filed].flatMap(([fingerprint, references]) =>
    [...references].map((issue) => ({ fingerprint, issue })),
  );
};
