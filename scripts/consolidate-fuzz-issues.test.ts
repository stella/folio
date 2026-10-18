import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  applyConsolidation,
  type ConsolidationStore,
  evidenceOf,
  identityOf,
  lanesOfWorkflow,
  planConsolidation,
  reproOf,
  pauseOnGitHubRefusal,
  summarize,
} from "./consolidate-fuzz-issues";
import {
  type CloseReason,
  type IssueContent,
  parseSignatureState,
  signatureOf,
} from "./fuzz-failure-issues";
import type { Issue } from "./fuzz-issue-classes";

const RUN = "https://github.com/stella/folio/actions/runs/";
const TEST = "consumer flow directNumbered / tracked-changes";
const REPRO = (seed: number) =>
  `bun scripts/consumer-scenarios.ts --seed ${seed} --steps 20 --runs 1 --collision-runs 0 --swarm 1 --only '^fuzz run 0 \\(' -- fuzz.test.ts`;

const fingerprintBody = (seeds: readonly number[], run: number, assertion: string) =>
  [
    `Run: ${RUN}${run}`,
    `Seeds: ${seeds.join(", ")}`,
    "Fingerprint: 7d8e8fd28bc8c8d8",
    `Test: ${TEST}`,
    `Assertion:\n\`\`\`\n${assertion}\n\`\`\``,
    ...seeds.map((seed) => `Seed: ${seed}\nRepro command:\n\`\`\`\n${REPRO(seed)}\n\`\`\``),
  ].join("\n\n");

const markdownBody = (seed: number, run: number, symptom: string) =>
  [
    `A fuzz flow failed in [continuous fuzz](${RUN}${run}), commit \`41984db4acc7\`.`,
    "",
    `**Test:** ${TEST}`,
    `**Symptom:** \`${symptom}\``,
    "**Minimized operations:** `deleteBlock+setBlockParagraphProperties`",
    "**Fingerprint:** `35240f905441125f`",
    `**Seed:** \`${seed}\``,
    "",
    "### Replay",
    "```sh",
    REPRO(seed),
    "```",
  ].join("\n");

const classBody = (seeds: readonly number[], run: number) =>
  [
    '`["consumer-flow","Missing numbering format for displayed label"]`',
    "",
    "| Fingerprint | Fixture | Flow | Mode | Seed | Path | One-line repro | First seen run |",
    "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ...seeds.map(
      (seed) =>
        `| 007e000b6c897c7d |  | deleteBlock+mergeBlockWithNext | tracked-changes | ${seed} |  | bun scripts/consumer-scenarios.ts --flow '{"a":"b\\|c"}' | ${RUN}${run} |`,
    ),
  ].join("\n");

const open = (number: number, title: string, body: string): Issue => ({
  number,
  title,
  body,
  state: "open",
  closedAt: null,
});

test("apply resumes from its checkpoint without repeating completed writes", async () => {
  const issues = fixture();
  const fake = fakeStore(issues);
  const plan = await planConsolidation(fake.store);
  const directory = mkdtempSync(path.join(tmpdir(), "fuzz-consolidation-"));
  const checkpoint = path.join(directory, "checkpoint.json");
  let calls = 0;
  const interrupted = {
    edit: async (number: number, content: IssueContent) => {
      calls += 1;
      if (calls === 2) throw new Error("simulated interruption");
      await fake.store.edit(number, content);
    },
    close: fake.store.close,
  };
  await expect(applyConsolidation(plan, interrupted, checkpoint)).rejects.toThrow(
    "simulated interruption",
  );
  const completedCalls = calls;
  await applyConsolidation(plan, interrupted, checkpoint);
  expect(calls).toBeGreaterThan(completedCalls);
  expect(fake.edits.filter(({ number }) => number === plan.groups.at(0)?.keep.number)).toHaveLength(
    1,
  );
  rmSync(directory, { recursive: true });
});

test("a GitHub refusal becomes a typed pause", () => {
  expect(() => pauseOnGitHubRefusal(new Error("HTTP 429 secondary rate limit"))).toThrow(
    "HTTP 429",
  );
});

const fakeStore = (issues: Issue[], comments: Record<number, string[]> = {}) => {
  const edits: { number: number; content: IssueContent }[] = [];
  const closes: { number: number; reason: CloseReason }[] = [];
  const store = {
    list: () => Promise.resolve(issues),
    comments: (number) => Promise.resolve(comments[number] ?? []),
    lanesOfRun: (runUrl) =>
      Promise.resolve(runUrl.endsWith("/9") ? ["continuous-fuzz"] : ["nightly-swarm-fuzz"]),
    edit: (number, content) => {
      edits.push({ number, content });
      const index = issues.findIndex((issue) => issue.number === number);
      const issue = issues.at(index);
      if (issue === undefined) throw new Error(`Unknown issue #${number}`);
      issues[index] = { ...issue, ...content };
      return Promise.resolve();
    },
    close: (number, reason) => {
      closes.push({ number, reason });
      const index = issues.findIndex((issue) => issue.number === number);
      const issue = issues.at(index);
      if (issue === undefined) throw new Error(`Unknown issue #${number}`);
      issues[index] = {
        number,
        title: issue.title,
        body: issue.body,
        state: "closed",
        closedAt: "2026-10-17",
        disposition: { type: "duplicate" },
      };
      return Promise.resolve();
    },
  } satisfies ConsolidationStore;
  return { store, edits, closes };
};

const fixture = () => [
  open(
    1863,
    `Fuzz failure [7d8e8fd28bc8c8d8]: ${TEST}`,
    fingerprintBody([740097491], 30, "Missing numbering format for displayed label (<n>)"),
  ),
  open(
    1746,
    `Fuzz failure [87ac1418886c9a57]: ${TEST}`,
    fingerprintBody([1809930746, 1809931056], 10, "Missing numbering format for displayed label •"),
  ),
  open(
    1410,
    `Fuzz failure [35240f905441125f]: ${TEST}: Missing numbering format for displayed label <n>.`,
    markdownBody(1576446700, 9, "Missing numbering format for displayed label <n>."),
  ),
  open(
    1415,
    `Fuzz failure [4f805f941a853afa]: ${TEST}: Missing numbering format for displayed label •`,
    classBody([1026397300, 1026398879], 8),
  ),
  open(
    1874,
    `Fuzz failure [e4869587aff7a192]: ${TEST}`,
    fingerprintBody([740097319], 30, "accept-all document differs after save"),
  ),
  open(1900, "Unrelated report", "No structured evidence"),
];

test("legacy bodies yield their identity, seeds and repro", () => {
  const [fingerprint, , markdown, table] = fixture();
  if (!fingerprint || !markdown || !table) throw new Error("Missing fixture");
  expect(identityOf(fingerprint)).toEqual({
    test: TEST,
    assertion: "Missing numbering format for displayed label (<n>)",
  });
  expect(identityOf(markdown)?.assertion).toBe("Missing numbering format for displayed label <n>.");
  expect(identityOf(table)?.assertion).toBe("Missing numbering format for displayed label •");
  expect(evidenceOf(fingerprint.body ?? "")).toEqual({ seeds: [740097491], occurrences: 1 });
  expect(evidenceOf("Seeds: 1, 2\n\n3 more seeds; see run artifacts.")).toEqual({
    seeds: [1, 2],
    occurrences: 5,
  });
  expect(evidenceOf(table.body ?? "").seeds).toEqual([1026397300, 1026398879]);
  expect(reproOf(markdown.body ?? "")).toEqual({ seed: 1576446700, command: REPRO(1576446700) });
  expect(reproOf(table.body ?? "")?.command).toBe(
    `bun scripts/consumer-scenarios.ts --flow '{"a":"b|c"}'`,
  );
});

test("consolidation keeps the oldest issue per signature and closes the rest as duplicates", async () => {
  const issues = fixture();
  const comments = { 1746: [`Run: ${RUN}20\n\nSeeds: 1809931334\n\nFingerprint: x`] };
  const fake = fakeStore(issues, comments);
  const plan = await planConsolidation(fake.store);
  expect(plan.open).toBe(6);
  expect(plan.before).toBe(6);
  expect(plan.groups).toHaveLength(2);
  expect(plan.unparsed.map(({ number }) => number)).toEqual([1900]);
  const [numbering, acceptAll] = plan.groups;
  expect(numbering?.keep.number).toBe(1410);
  expect(numbering?.duplicates.map(({ number }) => number)).toEqual([1415, 1746, 1863]);
  expect(numbering?.signature).toEqual(
    signatureOf({ test: TEST, assertion: "Missing numbering format for displayed label" }),
  );
  expect(acceptAll?.keep.number).toBe(1874);
  expect(acceptAll?.duplicates).toEqual([]);
  expect(fake.edits).toEqual([]);
  expect(fake.closes).toEqual([]);
  expect(summarize(plan)).toContain("Signatures after: 2");

  await applyConsolidation(plan, fake.store);
  expect(fake.closes).toEqual([
    { number: 1415, reason: "duplicate" },
    { number: 1746, reason: "duplicate" },
    { number: 1863, reason: "duplicate" },
  ]);
  expect(fake.edits.map(({ number }) => number)).toEqual([1410, 1874]);
  const kept = issues.find(({ number }) => number === 1410);
  const state = kept === undefined ? null : parseSignatureState(kept);
  expect(state?.seeds).toEqual([
    1576446700, 1026397300, 1026398879, 1809930746, 1809931056, 1809931334, 740097491,
  ]);
  expect(state?.occurrences).toBe(7);
  expect(state?.repro).toEqual({ seed: 740097491, command: REPRO(740097491) });
  expect(kept?.body).not.toContain("FOLIO_SCENARIO_");
  expect(state?.assertion).toBe("Missing numbering format for displayed label (<n>)");
  expect(state?.lanes.map(({ lane, cleanRuns }) => [lane, cleanRuns])).toEqual([
    ["nightly-swarm-fuzz", 0],
    ["continuous-fuzz", 0],
  ]);
  expect(state?.lastSeen).toEqual({ lane: "nightly-swarm-fuzz", run: `${RUN}30` });
  expect(issues.find(({ number }) => number === 1900)?.state).toBe("open");

  const rerun = fakeStore(issues);
  await applyConsolidation(await planConsolidation(rerun.store), rerun.store);
  expect(rerun.edits).toEqual([]);
  expect(rerun.closes).toEqual([]);
});

test("lanes come from the reporter jobs and inline filer runs of a workflow", () => {
  const workflows = path.resolve(import.meta.dir, "../.github/workflows");
  const read = (file: string) => readFileSync(path.join(workflows, file), "utf8");
  expect(lanesOfWorkflow(read("nightly-swarm-fuzz.yml"))).toEqual(["nightly-swarm-fuzz"]);
  expect(lanesOfWorkflow(read("continuous-fuzz.yml"))).toEqual(["continuous-fuzz"]);
  expect(lanesOfWorkflow(read("nightly-browser-input-fuzzer.yml"))).toHaveLength(4);
});
