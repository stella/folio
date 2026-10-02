import { describe, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import {
  failureMarker,
  failureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import type { Context } from "./fuzz-failure-issues";
import type { Issue, IssueStore } from "./fuzz-issue-classes";

import { hash32 } from "../test/commit-seed";
import {
  consumerSeed,
  judge,
  packageOf,
  propertyReplays,
  fileMainFindings,
  replayOutcome,
  readReplayAttempts,
  type ReplayAttempt,
} from "./fuzz-ejection";

const property = (fields: Record<string, unknown>): string =>
  `2026-09-29T10:00:00Z PROPERTY_FAILURE ${JSON.stringify({
    file: "packages/core/src/ai-edits/batchOverlap.test.ts",
    test: "refuses each conflict and applies the rest as one at a time would",
    seed: -440111207,
    path: "136:9:8:8",
    numRunsFactor: 1,
    ...fields,
  })}`;

describe("merge group fuzz replay", () => {
  test("the base replays the consumer flows under the seed the group's run derived", () => {
    const sha = "0302a42f5661aaaaaaaaaaaaaaaaaaaaaaaaaaaa";
    // scripts/consumer-scenarios.ts: commitSeed("consumer-scenarios fuzz") >>> 1.
    expect(consumerSeed(sha)).toBe(hash32(`${sha}\0consumer-scenarios fuzz`) >>> 1);
    expect(consumerSeed(sha)).toBeGreaterThanOrEqual(0);
  });

  test("property replays come from the failure lines, and only well-formed ones", () => {
    const log = [
      property({}),
      property({ path: "" }),
      property({ file: "../../etc/passwd.test.ts" }),
      property({ file: "packages/core/src/x.ts" }),
      property({ seed: "1; rm -rf /" }),
      property({ path: "1:2; echo" }),
      "PROPERTY_FAILURE {not json",
    ].join("\n");
    expect(propertyReplays(log)).toEqual([
      {
        file: "packages/core/src/ai-edits/batchOverlap.test.ts",
        title: "refuses each conflict and applies the rest as one at a time would",
        seed: -440111207,
        path: "136:9:8:8",
        factor: 1,
      },
      {
        file: "packages/core/src/ai-edits/batchOverlap.test.ts",
        title: "refuses each conflict and applies the rest as one at a time would",
        seed: -440111207,
        path: null,
        factor: 1,
      },
    ]);
    expect(packageOf("packages/core/src/ai-edits/batchOverlap.test.ts")).toEqual({
      dir: "packages/core",
      file: "src/ai-edits/batchOverlap.test.ts",
    });
    expect(packageOf("scripts/x.test.ts")).toEqual({ dir: ".", file: "scripts/x.test.ts" });
  });

  test("a group is pre-existing only when the base fails every way the group did", () => {
    const group = [
      { fingerprint: "aaaaaaaaaaaaaaaa", test: "consumer flow a / direct" },
      { fingerprint: "bbbbbbbbbbbbbbbb", test: "consumer flow b / direct" },
    ];
    expect(judge(group, new Set(["aaaaaaaaaaaaaaaa", "bbbbbbbbbbbbbbbb"]))).toMatchObject({
      verdict: "pre-existing",
      requeue: true,
    });
    expect(judge(group, new Set(["aaaaaaaaaaaaaaaa"]))).toMatchObject({
      verdict: "mixed",
      requeue: false,
    });
    expect(judge(group, new Set())).toMatchObject({ verdict: "introduced", requeue: false });
    expect(judge([], new Set(["aaaaaaaaaaaaaaaa"]))).toMatchObject({
      verdict: "none",
      requeue: false,
    });
  });
});

const mainSha = "1111111111111111111111111111111111111111";
const groupSha = "2222222222222222222222222222222222222222";
const context = {
  runUrl: "https://github.com/stella/folio/actions/runs/42",
  sha: mainSha,
  source: "merge group reproduced on main",
  date: "2026-10-02",
} satisfies Context;
const finding = (fixture: string) => {
  const error = new Error("step 3: no comment");
  const marker = failureMarker({
    test: `consumer flow ${fixture} / direct`,
    seed: 1026398907,
    path: null,
    repro: "FOLIO_SCENARIO_SEED=1026398907 bun scripts/consumer-scenarios.ts",
    failure: error,
  });
  const record = failureRecord(marker, error);
  return { record, seeds: [marker.seed], records: [record] };
};
const withStore = async (
  body: (fixture: {
    root: string;
    store: IssueStore;
    writes: Issue[];
    reads: () => number;
  }) => Promise<void>,
) => {
  const root = mkdtempSync(path.join(tmpdir(), "folio-main-replay-"));
  mkdirSync(path.join(root, "test"));
  writeFileSync(
    path.join(root, "test/known-failure-fingerprints.json"),
    JSON.stringify({ known: [] }),
  );
  const writes: Issue[] = [];
  let reads = 0;
  const store = {
    list: async () => {
      reads++;
      return writes;
    },
    create: async (title: string, issueBody: string) => {
      const issue = {
        number: writes.length + 1,
        title,
        body: issueBody,
        state: "open",
        closedAt: null,
      } satisfies Issue;
      writes.push(issue);
      return issue;
    },
    edit: async () => {
      throw new Error("Unexpected issue edit");
    },
    reopen: async () => {
      throw new Error("Unexpected issue reopen");
    },
  } satisfies IssueStore;
  try {
    await body({ root, store, writes, reads: () => reads });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
};

test("group-only failures file nothing through the issue store", async () => {
  await withStore(async ({ root, store, writes, reads }) => {
    const group = finding("stories");
    const judged = await fileMainFindings({
      findings: [group],
      attempts: [
        {
          requested: [group.record.marker.fingerprint],
          outcome: { status: "passed", fingerprints: [] },
        },
      ],
      context,
      root,
      store,
    });
    expect(judged.verdict).toBe("introduced");
    expect(judged.fingerprints.at(0)?.issue).toBeNull();
    expect(writes).toHaveLength(0);
    expect(reads()).toBe(0);
  });
});

test("matching main failures file exactly one class issue with only reproduced rows", async () => {
  await withStore(async ({ root, store, writes }) => {
    const findings = [finding("stories"), finding("comments"), finding("lists")];
    const confirmed = findings.slice(0, 2).map(({ record }) => record.marker.fingerprint);
    const judged = await fileMainFindings({
      findings,
      attempts: [
        {
          requested: findings.map(({ record }) => record.marker.fingerprint),
          outcome: { status: "failed", fingerprints: confirmed },
        },
      ],
      context,
      root,
      store,
    });
    expect(judged.verdict).toBe("mixed");
    expect(writes).toHaveLength(1);
    const body = writes.at(0)?.body;
    for (const fingerprint of confirmed) expect(body).toContain(fingerprint);
    expect(body).not.toContain(findings.at(2)?.record.marker.fingerprint);
    expect(body).toContain(context.runUrl);
    expect(judged.fingerprints.map(({ issue }) => issue)).toEqual(["#1", "#1", null]);
  });
});

test("unavailable or incomplete main replay prevents every store access", async () => {
  await withStore(async ({ root, store, writes, reads }) => {
    const findings = [finding("stories"), finding("comments")];
    const fingerprint = findings.at(0)?.record.marker.fingerprint;
    if (fingerprint === undefined) throw new Error("Missing fixture fingerprint");
    const valid = {
      requested: [fingerprint],
      outcome: { status: "failed", fingerprints: [fingerprint] },
    } satisfies ReplayAttempt;
    const cases: ReplayAttempt[][] = [
      [],
      [valid],
      [
        valid,
        { requested: [], outcome: { status: "unavailable", message: "Main replay could not run" } },
      ],
      [
        {
          requested: findings.map(({ record }) => record.marker.fingerprint),
          outcome: { status: "failed", fingerprints: [] },
        },
      ],
    ];
    for (const attempts of cases) {
      await expect(
        fileMainFindings({ findings, attempts, context, root, store }),
      ).rejects.toThrow();
      expect(writes).toHaveLength(0);
      expect(reads()).toBe(0);
    }
  });
});

test("replay execution distinguishes actual tests from setup, no-match, timeout and unrelated failure", () => {
  const marker = finding("stories").record.marker;
  const failure = `FOLIO_FAILURE ${JSON.stringify(marker)}\nℹ pass 0\nℹ fail 1`;
  expect(replayOutcome({ output: failure, exitCode: 1 })).toEqual({
    status: "failed",
    fingerprints: [marker.fingerprint],
  });
  for (const output of ["1 pass\n0 fail", "ℹ pass 1\nℹ fail 0"]) {
    expect(replayOutcome({ output, exitCode: 0 })).toEqual({ status: "passed", fingerprints: [] });
  }
  for (const input of [
    { output: "Module missing", exitCode: 1 },
    { output: "ℹ tests 0\nℹ pass 0\nℹ fail 0", exitCode: 0 },
    { output: "1 fail", exitCode: 1 },
    { output: failure, exitCode: null, error: "Replay timed out" },
    { output: failure, exitCode: 0 },
  ])
    expect(replayOutcome(input).status).toBe("unavailable");
});

test("report evidence requires a completed replay tied to the pinned main and group", async () => {
  await withStore(async ({ root }) => {
    const options = { baseLogs: root, groupSha, baseSha: mainSha };
    expect(() => readReplayAttempts(options)).toThrow("did not complete");
    const attempts = [
      {
        requested: [finding("stories").record.marker.fingerprint],
        outcome: { status: "passed", fingerprints: [] },
      },
    ];
    writeFileSync(
      path.join(root, "replay-complete.json"),
      JSON.stringify({ group: groupSha, main: mainSha, attempts }),
    );
    expect(readReplayAttempts(options)).toEqual(attempts);
    expect(() => readReplayAttempts({ ...options, baseSha: groupSha })).toThrow("does not match");
    const cli = Bun.spawnSync(
      [
        "bun",
        "scripts/fuzz-ejection.ts",
        "report",
        "--queue-logs",
        root,
        "--base-logs",
        path.join(root, "missing"),
        "--group-sha",
        groupSha,
        "--base-sha",
        mainSha,
        "--dry-run",
      ],
      { cwd: path.resolve(import.meta.dir, ".."), stdout: "pipe", stderr: "pipe" },
    );
    expect(cli.exitCode).not.toBe(0);
    expect(cli.stderr.toString()).toContain("Main replay did not complete");
  });
});

test("ejection workflow pins main and separates replay from issue writes", () => {
  const workflow = Bun.YAML.parse(
    readFileSync(
      path.resolve(import.meta.dir, "../.github/workflows/fuzz-ejection-replay.yml"),
      "utf8",
    ),
  );
  const { replay, report } = workflow.jobs;
  expect(replay.permissions).toEqual({ actions: "read", contents: "read" });
  expect(report.needs).toBe("replay");
  expect(report.concurrency).toEqual({ group: "fuzz-failure-issues", "cancel-in-progress": false });
  const pin = replay.steps.find((step: { id?: string }) => step.id === "base");
  expect(pin.run).toContain('commits/main" --jq');
  expect(pin.run).not.toContain(".parents[0]");
  const checkouts = replay.steps.filter((step: { uses?: string }) =>
    step.uses?.startsWith("actions/checkout@"),
  );
  expect(checkouts).toHaveLength(2);
  for (const checkout of checkouts) expect(checkout.with.ref).toBe("${{ steps.base.outputs.sha }}");
  const execution = replay.steps.find((step: { run?: string }) =>
    step.run?.includes("fuzz-ejection.ts replay"),
  );
  expect(execution.run).toContain('--base-sha "${BASE_SHA}"');
  expect(execution.env.BASE_SHA).toBe("${{ steps.base.outputs.sha }}");
  expect(report.steps.some((step: { run?: string }) => step.run?.includes("gh pr comment"))).toBe(
    false,
  );
});
