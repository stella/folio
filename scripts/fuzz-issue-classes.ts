import { Result } from "better-result";
import { classOfTitle, failureClass } from "./fuzz-failure-class";
import type { Context, Filed, Finding } from "./fuzz-failure-issues";

export type Issue = {
  number: number;
  title: string;
  state: "open" | "closed";
  body: string | null;
  closedAt: string | null;
};

export type IssueStore = {
  list: () => Promise<Issue[]>;
  create: (title: string, body: string) => Promise<Issue>;
  edit: (number: number, body: string) => Promise<void>;
  reopen: (number: number) => Promise<void>;
};

type SeedRow = {
  fingerprint: string;
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
const MAX_NEW_ISSUES = 15;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isSeedRow = (value: unknown): value is SeedRow =>
  isRecord(value) &&
  typeof value["fingerprint"] === "string" &&
  /^[0-9a-f]{16}$/u.test(value["fingerprint"]) &&
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

const readClassState = (body: string | null, key: string): ClassState | null => {
  const json = body?.match(/<!-- fuzz-class-state (\{[^\n]*\}) -->/u)?.at(1);
  if (json === undefined) return null;
  const parsed = Result.try((): unknown => JSON.parse(json));
  if (parsed.isErr()) return null;
  const value = parsed.value;
  if (
    !isRecord(value) ||
    value["key"] !== key ||
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
  if (marker === undefined) return undefined;
  const parsed = Result.try((): unknown => JSON.parse(marker));
  if (
    parsed.isErr() ||
    !Array.isArray(parsed.value) ||
    !parsed.value.every((value: unknown) => typeof value === "string")
  )
    return undefined;
  return JSON.stringify(parsed.value);
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

const latestIssue = (issues: readonly Issue[]): Issue | undefined => {
  const open = issues.filter(({ state }) => state === "open");
  if (open.length > 0) return open.toSorted((a, b) => b.number - a.number).at(0);
  return issues
    .toSorted(
      (a, b) =>
        (Date.parse(b.closedAt ?? "") || 0) - (Date.parse(a.closedAt ?? "") || 0) ||
        b.number - a.number,
    )
    .at(0);
};

const rowIdentity = ({ fingerprint, seed, path }: SeedRow): string =>
  JSON.stringify([fingerprint, seed, path]);
const tableCell = (value: string): string =>
  value.replaceAll("|", "\\|").replaceAll("\n", " ").replaceAll("`", "\\`");

type ClassBodyOptions = {
  signature: ReturnType<typeof failureClass>;
  state: ClassState;
  note: string | null;
  legacyReport: string | null;
};
const classBody = ({ signature, state, note, legacyReport }: ClassBodyOptions): string => {
  const rows = state.seeds.map(
    (row) =>
      `| ${row.fingerprint} | ${tableCell(row.flow)} | ${tableCell(row.mode)} | ${row.seed} | ${tableCell(row.path ?? "")} | ${tableCell(row.repro)} | ${tableCell(row.firstSeenRun)} |`,
  );
  const json = JSON.stringify(state).replaceAll("<", "\\u003c").replaceAll(">", "\\u003e");
  return [
    `Class signature: \`${signature.key}\``,
    `Failure: ${signature.message}`,
    `Seen in ${state.runs.length} distinct runs; ${state.seeds.length} replay rows.`,
    ...(note === null ? [] : [note]),
    "",
    "| Fingerprint | Flow | Mode | Seed | Path | One-line repro | First seen run |",
    "| --- | --- | --- | --- | --- | --- | --- |",
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

type ClassGroup = {
  signature: ReturnType<typeof failureClass>;
  rows: SeedRow[];
  fingerprints: Set<string>;
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
  const filed = new Map<string, Set<string | null>>();
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
      const aliases = issues.filter((issue) =>
        legacyIdentity(issue).some(
          (identity) =>
            identity === record.marker.fingerprint || identity === record.marker.primary,
        ),
      );
      const alias = latestIssue(aliases);
      const original = failureClass(record.marker.test, record.marker.assertion);
      const signature = (alias === undefined ? null : classOfTitle(alias.title)) ?? original;
      let group = groups.get(signature.key);
      if (group === undefined) {
        group = { signature, rows: [], fingerprints: new Set() };
        groups.set(signature.key, group);
      }
      group.fingerprints.add(fingerprint);
      const replays = record.replays.length > 0 ? record.replays : [record.marker.repro];
      group.rows.push({
        fingerprint: record.marker.fingerprint,
        flow: record.marker.flow ?? original.flow,
        mode: original.mode,
        seed: record.marker.seed,
        path: record.marker.path,
        repro: replays.at(0) ?? record.marker.repro,
        replays: [...replays],
        evidence: [
          record.error,
          record.marker.diff ?? "",
          record.flow === undefined ? "" : JSON.stringify(record.flow),
        ]
          .join("\n")
          .slice(0, 600),
        firstSeenRun: run,
      });
    }
  }
  let created = 0;
  for (const { signature, rows, fingerprints } of groups.values()) {
    const matching = issues.filter(
      (issue) =>
        markedKey(issue) === signature.key ||
        (markedKey(issue) === undefined && legacyKey(issue) === signature.key),
    );
    const previous = latestIssue(matching);
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
    let issue: Issue | undefined;
    if (reuse) {
      issue = previous;
      if (recent) {
        await store.reopen(issue.number);
        issue.state = "open";
        issue.closedAt = null;
      }
      if (changed || recent) {
        issue.body = classBody({ signature, state, note, legacyReport });
        await store.edit(issue.number, issue.body);
      }
    } else if (created < maxNewIssues) {
      const title = `${signature.area}: ${signature.message}`;
      issue = await store.create(
        title.length > 240 ? `${title.slice(0, 239)}…` : title,
        classBody({ signature, state, note, legacyReport }),
      );
      issues.push(issue);
      created += 1;
    }
    for (const fingerprint of fingerprints) {
      const references = filed.get(fingerprint) ?? new Set<string | null>();
      references.add(issue === undefined ? null : `#${issue.number}`);
      filed.set(fingerprint, references);
    }
  }
  return [...filed].flatMap(([fingerprint, references]) =>
    [...references].map((issue) => ({ fingerprint, issue })),
  );
};
