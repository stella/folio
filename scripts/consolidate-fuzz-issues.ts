#!/usr/bin/env bun

import { $ } from "bun";
import { TaggedError } from "better-result";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  type CloseReason,
  cliRepro,
  githubIssueStore,
  type IssueContent,
  mergeSeeds,
  parseSignatureState,
  renderSignatureIssue,
  type Signature,
  type SignatureState,
  signatureOf,
} from "./fuzz-failure-issues";
import type { Issue } from "./fuzz-issue-classes";

const ROOT = path.resolve(import.meta.dir, "..");
const REPORTER = "./.github/workflows/fuzz-failure-report.yml";
const CONSOLIDATED_LANE = "consolidated";
const UNAVAILABLE = "unavailable";
const RUN_URL = /https:\/\/github\.com\/[\w.-]+\/[\w.-]+\/actions\/runs\/(\d+)/gu;

export type ConsolidationStore = {
  list: () => Promise<Issue[]>;
  comments: (number: number) => Promise<string[]>;
  lanesOfRun: (runUrl: string) => Promise<string[]>;
  edit: (number: number, content: IssueContent) => Promise<void>;
  close: (number: number, reason: CloseReason) => Promise<void>;
};

type Identity = { test: string; assertion: string };
type Repro = SignatureState["repro"];
type Evidence = { seeds: number[]; occurrences: number };

const integers = (text: string): number[] =>
  text
    .split(",")
    .map((seed) => seed.trim())
    .filter((seed) => /^-?\d+$/u.test(seed))
    .map(Number);

const CLASS_ROW =
  /^\| [0-9a-f]{16} \| (?:\\.|[^|\\])* \| (?:\\.|[^|\\])* \| (?:\\.|[^|\\])* \| (-?\d+) \| (?:\\.|[^|\\])* \| ((?:\\.|[^|\\])*) \|/gmu;
const unescapeCell = (cell: string): string => cell.trim().replaceAll(/\\([|`\\])/gu, "$1");

export const identityOf = ({ title, body }: Pick<Issue, "title" | "body">): Identity | null => {
  const text = body ?? "";
  const test = /^Test: (.*)$/mu.exec(text)?.at(1) ?? /^\*\*Test:\*\* (.*)$/mu.exec(text)?.at(1);
  const assertion =
    /^Assertion:\n(`{3,})\n([\s\S]*?)\n\1$/mu.exec(text)?.at(2) ??
    /^\*\*Symptom:\*\* `(.*)`$/mu.exec(text)?.at(1);
  if (test !== undefined && assertion !== undefined) return { test, assertion };
  const [, titled, symptom] =
    /^Fuzz failure \[[0-9a-f]{16}\]: (consumer flow .+? \/ [^:]+): (.*)$/u.exec(title) ?? [];
  return titled === undefined || symptom === undefined
    ? null
    : { test: titled, assertion: symptom };
};

export const evidenceOf = (text: string): Evidence => {
  const lists = [...text.matchAll(/^Seeds: (.*)$/gmu)].flatMap(([, seeds = ""]) => integers(seeds));
  if (lists.length > 0) {
    const omitted = [...text.matchAll(/^(\d+) more seeds; see run artifacts\.$/gmu)].reduce(
      (sum, [, count = "0"]) => sum + Number(count),
      0,
    );
    return { seeds: lists, occurrences: lists.length + omitted };
  }
  const single = [...text.matchAll(/^\*\*Seed:\*\* `(-?\d+)`$/gmu)].map(([, seed]) => Number(seed));
  if (single.length > 0) return { seeds: single, occurrences: single.length };
  const rows = [...text.matchAll(CLASS_ROW)].map(([, seed]) => Number(seed));
  return { seeds: rows, occurrences: rows.length };
};

export const reproOf = (text: string): Repro | null => {
  const [, seed, , command] =
    /^Seed: (-?\d+)\nRepro command:\n(`{3,})\n([\s\S]*?)\n\2$/mu.exec(text) ?? [];
  if (seed !== undefined && command !== undefined)
    return { seed: Number(seed), command: cliRepro(command) };
  const replay = /^### Replay\n(`{3,})sh\n([\s\S]*?)\n\1$/mu.exec(text)?.at(2);
  const replaySeed = /^\*\*Seed:\*\* `(-?\d+)`$/mu.exec(text)?.at(1);
  if (replay !== undefined && replaySeed !== undefined) {
    return { seed: Number(replaySeed), command: cliRepro(replay) };
  }
  const [row] = [...text.matchAll(CLASS_ROW)];
  const [, rowSeed, rowRepro] = row ?? [];
  return rowSeed === undefined || rowRepro === undefined
    ? null
    : { seed: Number(rowSeed), command: cliRepro(unescapeCell(rowRepro)) };
};

const runsOf = (text: string): string[] => [...text.matchAll(RUN_URL)].map(([url]) => url);
const runNumber = (url: string): number => Number(/(\d+)$/u.exec(url)?.at(1) ?? 0);

type Member = { issue: Issue; texts: string[]; identity: Identity; state: SignatureState | null };

type Group = { signature: Signature; keep: Issue; duplicates: Issue[]; state: SignatureState };

export type ConsolidationPlan = {
  open: number;
  before: number;
  groups: Group[];
  unparsed: Issue[];
};

class ConsolidationError extends TaggedError("ConsolidationError")<{ message: string }> {}

type MergeOptions = { members: readonly Member[]; lanesOfRun: (run: string) => Promise<string[]> };
const mergeGroup = async ({ members, lanesOfRun }: MergeOptions): Promise<Group> => {
  const ordered = members.toSorted((a, b) => a.issue.number - b.issue.number);
  const [first, ...rest] = ordered;
  if (first === undefined) throw new ConsolidationError({ message: "Empty signature group" });
  if (rest.length === 0 && first.state !== null) {
    return {
      signature: first.state.signature,
      keep: first.issue,
      duplicates: [],
      state: first.state,
    };
  }
  const newest = ordered.toReversed();
  const signature = first.state?.signature ?? signatureOf(first.identity);
  let seeds: number[] = [];
  let occurrences = 0;
  const lanes = new Map<string, string>();
  for (const { texts, state } of ordered) {
    if (state !== null) {
      seeds = mergeSeeds(seeds, state.seeds);
      occurrences += state.occurrences;
      for (const { lane, lastRun } of state.lanes) lanes.set(lane, lastRun);
      continue;
    }
    for (const text of texts) {
      const evidence = evidenceOf(text);
      seeds = mergeSeeds(seeds, evidence.seeds);
      occurrences += evidence.occurrences;
    }
  }
  const runs = [
    ...new Set(
      ordered.flatMap(({ state, texts }) => (state === null ? texts.flatMap(runsOf) : [])),
    ),
  ].toSorted((a, b) => runNumber(a) - runNumber(b));
  for (const run of runs) for (const lane of await lanesOfRun(run)) lanes.set(lane, run);
  const latest = newest.at(0) ?? first;
  const latestRun = runs.at(-1);
  const latestLane = latestRun === undefined ? undefined : (await lanesOfRun(latestRun)).at(0);
  const lastSeen = latest.state?.lastSeen ?? {
    lane: latestLane ?? CONSOLIDATED_LANE,
    run: latestRun ?? UNAVAILABLE,
  };
  const repro = newest
    .flatMap(({ state, texts }) =>
      state === null ? texts.toReversed().map(reproOf) : [state.repro],
    )
    .find((candidate) => candidate !== null) ?? {
    seed: seeds.at(-1) ?? 0,
    command: UNAVAILABLE,
  };
  return {
    signature,
    keep: first.issue,
    duplicates: rest.map(({ issue }) => issue),
    state: {
      signature,
      occurrences,
      seeds,
      lastSeen,
      lanes: [...lanes].map(([lane, lastRun]) => ({ lane, lastRun, cleanRuns: 0 })),
      assertion: latest.state?.assertion ?? latest.identity.assertion,
      repro,
    },
  };
};

export const planConsolidation = async (store: ConsolidationStore): Promise<ConsolidationPlan> => {
  const open = (await store.list()).filter(({ state }) => state === "open");
  const members = new Map<string, Member[]>();
  const unparsed: Issue[] = [];
  for (const issue of open) {
    const state = parseSignatureState(issue);
    const identity =
      state === null
        ? identityOf(issue)
        : { test: state.signature.test, assertion: state.assertion };
    if (identity === null) {
      unparsed.push(issue);
      continue;
    }
    const texts = state === null ? [issue.body ?? "", ...(await store.comments(issue.number))] : [];
    const id = state?.signature.id ?? signatureOf(identity).id;
    const member = { issue, texts, identity, state };
    const group = members.get(id);
    if (group === undefined) members.set(id, [member]);
    else group.push(member);
  }
  const groups: Group[] = [];
  for (const group of members.values()) {
    groups.push(await mergeGroup({ members: group, lanesOfRun: store.lanesOfRun }));
  }
  const ids = new Set(
    open.map(
      ({ number, title }) => /^Fuzz failure \[([0-9a-f]{16})\]/u.exec(title)?.at(1) ?? `#${number}`,
    ),
  );
  return {
    open: open.length,
    before: ids.size,
    groups: groups.toSorted(
      (a, b) => b.duplicates.length - a.duplicates.length || a.keep.number - b.keep.number,
    ),
    unparsed,
  };
};

export const applyConsolidation = async (
  plan: ConsolidationPlan,
  store: Pick<ConsolidationStore, "edit" | "close">,
): Promise<void> => {
  for (const { keep, duplicates, state } of plan.groups) {
    const content = renderSignatureIssue(state);
    if (content.title !== keep.title || content.body !== keep.body)
      await store.edit(keep.number, content);
    for (const duplicate of duplicates) await store.close(duplicate.number, "duplicate");
  }
};

export const summarize = (plan: ConsolidationPlan): string =>
  [
    `Open fuzz issues: ${plan.open}`,
    `Signatures before: ${plan.before}`,
    `Signatures after: ${plan.groups.length}`,
    `Issues to close: ${plan.groups.reduce((sum, { duplicates }) => sum + duplicates.length, 0)}`,
    `Unparsed (left open): ${plan.unparsed.length === 0 ? "none" : plan.unparsed.map(({ number }) => `#${number}`).join(", ")}`,
    "",
    ...plan.groups.map(
      ({ keep, duplicates, state }) =>
        `#${keep.number} keeps ${state.signature.test}: ${state.signature.failure} (${state.occurrences} occurrences, ${state.seeds.length} seeds, lanes: ${state.lanes.map(({ lane }) => lane).join(", ") || "none"})${duplicates.length === 0 ? "" : `; closes ${duplicates.map(({ number }) => `#${number}`).join(", ")}`}`,
    ),
  ].join("\n");

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

export const lanesOfWorkflow = (source: string): string[] => {
  const workflow: unknown = Bun.YAML.parse(source);
  if (!isRecord(workflow) || !isRecord(workflow["jobs"])) return [];
  return Object.values(workflow["jobs"]).flatMap((job) => {
    if (!isRecord(job)) return [];
    const inputs = isRecord(job["with"]) ? job["with"] : {};
    if (job["uses"] === REPORTER) return typeof inputs["lane"] === "string" ? [inputs["lane"]] : [];
    const steps = Array.isArray(job["steps"]) ? job["steps"] : [];
    return steps.flatMap((step) =>
      isRecord(step) && typeof step["run"] === "string"
        ? [...step["run"].matchAll(/--lane (\S+)/gu)].map(([, lane = ""]) => lane)
        : [],
    );
  });
};

const githubConsolidationStore = (): ConsolidationStore => {
  const issues = githubIssueStore();
  const cache = new Map<string, string[]>();
  return {
    list: issues.list,
    comments: issues.comments,
    edit: issues.edit,
    close: issues.close,
    lanesOfRun: async (runUrl) => {
      const cached = cache.get(runUrl);
      if (cached !== undefined) return cached;
      const id = runNumber(runUrl);
      const [file = "", event = ""] = (
        await $`gh api ${`repos/{owner}/{repo}/actions/runs/${id}`} --jq ${"[.path, .event] | @tsv"}`.text()
      )
        .trim()
        .split("\t");
      const workflow = path.join(ROOT, file);
      const lanes =
        event === "schedule" && existsSync(workflow)
          ? lanesOfWorkflow(readFileSync(workflow, "utf8"))
          : [];
      const resolved = lanes.length === 1 ? lanes : [];
      cache.set(runUrl, resolved);
      return resolved;
    },
  };
};

const main = async (): Promise<void> => {
  const args = process.argv.slice(2);
  const unknown = args.filter((arg) => arg !== "--apply");
  if (unknown.length > 0) {
    throw new ConsolidationError({ message: `Unknown arguments: ${unknown.join(" ")}` });
  }
  const store = githubConsolidationStore();
  const plan = await planConsolidation(store);
  console.log(summarize(plan));
  if (!args.includes("--apply")) {
    console.log("\nDry run: nothing changed. Pass --apply to edit and close issues.");
    return;
  }
  await applyConsolidation(plan, store);
};

if (import.meta.main) {
  await main();
}
